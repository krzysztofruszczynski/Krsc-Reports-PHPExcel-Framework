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
     * @var String
     */
    protected $_sStyleKey = KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
    
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
    
    /**
     * 
     * @param String $sStyleKey
     * @param Boolean $bIfStyleNotExistsSelectDefault
     * @return Boolean true if style was found, false otherwise
     */
    public function setStyleKey( $sStyleKey, $bIfStyleNotExistsSelectDefault = false )
    {
       $bIsValid = $this->_oStyle->isValidStyleKey( $sStyleKey );
       
       if( $bIsValid )
       {   // style key exists 
           $this->_sStyleKey = $sStyleKey; 
       }
       else if( $bIfStyleNotExistsSelectDefault )
       {   // style key not exists but user wanted in such situation to switch to default style (otherwise - previously set style is used) 
           $this->_oStyle = KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
       }
       
       return $bIsValid;
    }
    
    public function constructCell( $iColumnId, $iRowId )
    {
        //die( var_dump( $this->_oStyle->getStyleArray() ) );
        
        // setting styles
        self::$_oPHPExcel->getActiveSheet()->getStyleByColumnAndRow( $iColumnId, $iRowId )->applyFromArray( $this->_oStyle->getStyleArray( $this->_sStyleKey ) );
        
        // constructing phpexcel element
        self::$_oPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $iColumnId, $iRowId, $this->_mValue );
    }
    
}