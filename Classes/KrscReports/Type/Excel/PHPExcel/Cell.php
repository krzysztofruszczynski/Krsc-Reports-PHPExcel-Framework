<?php
/**
 * Class which one object represents one cell. It applies styles, value and type of cell.
 * At the end cell with given coordinates and properties can be constructed (methods inside class
 * use directly PHPExcel commands). 
 * 
 * @category KrscReports
 * @package KrscReports_Type
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Cell
{
    /**
     * @var PHPExcel instance of used PHPExcel object
     */
    protected static $_oPHPExcel;
    
    /**
     * @var mixed value of current cell
     */
    protected $_mValue;
    
    /**
     * @var string type of current cell 
     */
    protected $_sType;
    
    /**
     * @var KrscReports_Type_Excel_PHPExcel_Style object for managing style collections
     */
    protected $_oStyle;
    
    /**
     * @var string currently selected key of style collection
     */
    protected $_sStyleKey = KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT;
    
    /**
     * Constructor which updates PHPExcel instance inside a class.
     * @return Void
     */
    public function __construct()
    {
        self::$_oPHPExcel = KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject();
    }
    
    /**
     * Method for setting object for managing style collections.
     * @param KrscReports_Type_Excel_PHPExcel_Style $oStyle instance of object for managing style collections
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setStyleObject( KrscReports_Type_Excel_PHPExcel_Style $oStyle )
    {
        $this->_oStyle = $oStyle;
        return $this;
    }
    
    /**
     * Method for setting value of current cell.
     * @param mixed $mValue value of current cell (can be text, numeric or formula: starts with =)
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setValue( $mValue )
    {
        $this->_mValue = $mValue;
        return $this;
    }
    
    /**
     * Method setting type of cell.
     * @param string $sType type of cell to be set
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function setType( $sType )
    {
        $this->_sType = $sType;
        return $this;
    }
    
    /**
     * Method for setting current style key.
     * @param string $sStyleKey style key to be set (key under which style collection is stored)
     * @param boolean $bIfStyleNotExistsSelectDefault if true and style key was not found, default style key is set 
     * (if it is like by default false, previously set style key remains)
     * @return boolean true if style was found, false otherwise
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
    
    /**
     * Method returns letter coordinate for numeric coordinate of column.
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @return string letter coordinate of given column
     */
    public function getColumnDimension( $iColumnId )
    {
        return self::$_oPHPExcel->getActiveSheet()->getColumnDimensionByColumn( $iColumnId )->getColumnIndex();
    }
    
    /**
     * Method creates cell with previously set properties on given in method's input coordinates 
     * (with one object with set properties more than one cell can be created).
     * @param integer $iColumnId numeric coordinate of column (starts from 0)
     * @param integer $iRowId numeric coordinate of row (starts from 1)
     * @return KrscReports_Type_Excel_PHPExcel_Cell object on which method was executed
     */
    public function constructCell( $iColumnId, $iRowId )
    {        
        // setting styles
        self::$_oPHPExcel->getActiveSheet()->getStyleByColumnAndRow( $iColumnId, $iRowId )->applyFromArray( $this->_oStyle->getStyleArray( $this->_sStyleKey ) );
        
        // constructing phpexcel element
        self::$_oPHPExcel->getActiveSheet()->setCellValueByColumnAndRow( $iColumnId, $iRowId, $this->_mValue );        
        
        return $this;
    }
    
}