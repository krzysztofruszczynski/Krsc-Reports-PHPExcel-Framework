<?php
/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Type_Excel_PHPExcel_Style_Default
{
    /**
     * key for element, which is created by this object
     */
    const ARRAY_KEY = '';    
    
    const KEY_COLOR_RGB = 'rgb';
    
    abstract public function getStyleArray();
    
    public function getArrayKey()
    {
        return static::ARRAY_KEY;
    }
    
    protected static function _attachToArray( $aOutput, $sKey, $aStyle )
    {
        if( !empty( $aStyle ) )
        {
            $aOutput[$sKey] = $aStyle;
        }
        
        return $aOutput;
    }
    
    protected static function _getColorArray( $sColor )
    {
        return array( self::KEY_COLOR_RGB => $sColor );
    }
}
