<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Type_Excel_PHPExcel_Style
{
    const KEY_DEFAULT = 'default';
    
    /**
     * @var Array array of KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator
     */
    protected $_aStyleIterator;
    
    public function setStyleCollection( KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection $oStyleCollection, $sKey = self::KEY_DEFAULT )
    {
        $this->_aStyleIterator[$sKey] = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Iterator( $oStyleCollection );
    }
    
    /**
     * 
     * @param String $sKey key to verify
     * @return Boolean true if style key exists, false otherwise
     */
    public function isValidStyleKey( $sKey )
    {
        return isset( $this->_aStyleIterator[$sKey] ) ? true : false;
    }
    
    public function getStyleArray( $sKey = self::KEY_DEFAULT ) 
    {
        if( !isset( $this->_aStyleIterator[$sKey] ) )
        {   // requested key not found - switching to default
            $sKey = self::KEY_DEFAULT;
            if( !isset( $this->_aStyleIterator[$sKey] ) )
            {   // default style not set - returning empty style array
                return array();
            }
        }
        
        $this->_aStyleIterator[$sKey]->resetIterator();
        $aOutput = array();
        
        while( $this->_aStyleIterator[$sKey]->hasNextElement() )
        {
            $oStyleElement = $this->_aStyleIterator[$sKey]->getStyleElement();
            $aOutput[$oStyleElement->getArrayKey()] = $oStyleElement->getStyleArray();
        }
        
        return $aOutput;
    }
}
