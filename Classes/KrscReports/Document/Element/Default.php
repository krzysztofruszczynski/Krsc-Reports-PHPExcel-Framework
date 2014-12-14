<?php

/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element_Default extends KrscReports_Document_Element
{
    /**
     * @var KrscReports_Builder_Abstract
     */
    protected $_oBuilder;
    
    public function setBuilder( $oBuilder )
    {
        $this->_oBuilder = $oBuilder;
        return $this;
    }
    
    /**
     * TO BE CALLED ONLY FROM KrscReports_Document_Element
     * @param String $sGroupName
     * @return KrscReports_Document_Element_Default
     */
    public function setInnerGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        $this->_oBuilder->setGroupName( $sGroupName );
        return $this;
    }
    
    public function beforeConstructDocument()
    {
        $this->_oBuilder->setStartWidth( ( isset( $this->_aActualWidths[$this->_sGroupName] ) ? $this->_aActualWidths[$this->_sGroupName] : 0 ) );
        $this->_oBuilder->setStartHeight( ( isset( $this->_aActualHeights[$this->_sGroupName] ) ? $this->_aActualHeights[$this->_sGroupName] : 1 ), false );
        return $this;
    }
    
    /**
     * 
     * @param Integer $iStartWidth (starts from 0)
     */
    public function setStartWidth( $iStartWidth )
    {
        $this->_aActualWidths[$this->_sGroupName] = $iStartWidth;
    }
    
    /**
     * 
     * @param Integer $iStartHeight (starts with 1)
     */
    public function setStartHeight( $iStartHeight, $bForceHeight = true )
    {
        $this->_aActualHeights[$this->_sGroupName] = $iStartHeight;
        
        if( $bForceHeight )
        {   // height of element is set manually
            $this->_oBuilder->setStartHeight( $iStartHeight );
        }
    }
    
    public function getActualHeight()
    {
        return $this->_oBuilder->getActualHeight();
    }
    
    public function afterConstructDocument()
    {   
        $this->_aActualHeights[$this->_sGroupName] = $this->_oBuilder->getActualHeight();
        return $this;
    }
    
    public function constructDocument() 
    {
        $this->_oBuilder->constructDocument();
        return $this;
    }
}
