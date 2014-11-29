<?php

/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
abstract class KrscReports_Builder_Abstract
{
    /**
     * @var String name of group in which this content is generated 
     */
    protected $_sGroupName = '';
    
    protected $_iActualWidth = 0;
    
    protected $_iActualHeight = 0;
    
    /**
     * @var Array 
     */
    protected $_aData;
    
    public function setStartWidth( $iStartWidth )
    {
        $this->_iActualWidth = $iStartWidth;
    }
    
    public function setStartHeight( $iStartHeight )
    {
        $this->_iActualHeight = $iStartHeight;
    }
    
    /**
     * 
     * @param Array $aData
     * @return KrscReports_Document_Element_Default
     */
    public function setData( $aData )
    {
        $this->_aData = $aData;
        return $this;
        
    }
    
    public function getActualWidth()
    {
        return $this->_iActualWidth;
    }
    
    public function getActualHeight()
    {
        return $this->_iActualHeight;
    }
    
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
    }
    
    abstract public function constructDocument();
}
?>
