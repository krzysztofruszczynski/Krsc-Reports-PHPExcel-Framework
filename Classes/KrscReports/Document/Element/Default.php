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
    }
    
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        $this->_oBuilder->setGroupName( $sGroupName );
    }
    
    protected function beforeConstructDocument()
    {
        $this->_oBuilder->setStartWidth( ( isset( $this->_aActualWidths[$this->_sGroupName] ) ? $this->_aActualWidths[$this->_sGroupName] : 0 ) );
        $this->_oBuilder->setStartHeight( ( isset( $this->_aActualHeights[$this->_sGroupName] ) ? $this->_aActualHeights[$this->_sGroupName] : 0 ) );
    }
    
    public function afterConstructDocument()
    {
        $this->_aActualWidths[$this->_sGroupName] = $this->_oBuilder->getActualWidth();
        $this->_aActualHeights[$this->_sGroupName] = $this->_oBuilder->getActualHeight();
    }
    
    public function constructDocument() 
    {
        $this->_oBuilder->constructDocument();
    }
}
