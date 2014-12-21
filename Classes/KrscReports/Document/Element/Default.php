<?php

/**
 * Class for single elements (parts which can be placed in spreadsheets) to inherit from
 * (manages with setting builder, group name, width and heights).
 * 
 * @category KrscReports
 * @package KrscReports_Document
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element_Default extends KrscReports_Document_Element
{
    /**
     * @var KrscReports_Builder_Abstract object of builder responsible for creating element
     */
    protected $_oBuilder;
    
    /**
     * Method sets builder for the current element.
     * @param KrscReports_Builder_Abstract $oBuilder builder to be set
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function setBuilder( KrscReports_Builder_Abstract $oBuilder )
    {
        $this->_oBuilder = $oBuilder;
        return $this;
    }
    
    /**
     * Method sets group name inside element and builder. TO BE CALLED ONLY FROM KrscReports_Document_Element.
     * @param string $sGroupName name of group, to which current element belongs to
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function setInnerGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        $this->_oBuilder->setGroupName( $sGroupName );
        return $this;
    }
    
    /**
     * Code which is executed before creating document (setting width and height)
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function beforeConstructDocument()
    {
        $this->_oBuilder->setStartWidth( ( isset( $this->_aActualWidths[$this->_sGroupName] ) ? $this->_aActualWidths[$this->_sGroupName] : 0 ) );
        $this->_oBuilder->setStartHeight( ( isset( $this->_aActualHeights[$this->_sGroupName] ) ? $this->_aActualHeights[$this->_sGroupName] : 1 ), false );
        return $this;
    }
    
    /**
     * Method for setting start width for the current element (1 number = 1 cell).
     * If not set, by default it is 0 (starts from the beginning). Method can be invoked by user.
     * @param integer $iStartWidth width, at which element starts (starts from 0)
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function setStartWidth( $iStartWidth )
    {
        $this->_aActualWidths[$this->_sGroupName] = $iStartWidth;
        return $this;
    }
    
    /**
     * Method for setting start height for the current element (1 number = 1 row). Invoked by KrscReports_Document_Element composite
     * (automatically calculating height) or by user (than height is set manually and would not be changed later).
     * @param integer $iStartHeight start height (starts with 1)
     * @param boolean $bForceHeight (default: true) if true, given height is set also on builder what means that it would not be changed
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function setStartHeight( $iStartHeight, $bForceHeight = true )
    {
        $this->_aActualHeights[$this->_sGroupName] = $iStartHeight;
        
        if( $bForceHeight )
        {   // height of element is set manually
            $this->_oBuilder->setStartHeight( $iStartHeight );
        }
        return $this;
    }
    
    /**
     * Method returns current height given by builder. 
     * @return integer actual height (starts from 1)
     */
    public function getActualHeight()
    {
        return $this->_oBuilder->getActualHeight();
    }
    
    /**
     * Code which is executed after creating document (updating actual height).
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function afterConstructDocument()
    {   
        $this->_aActualHeights[$this->_sGroupName] = $this->_oBuilder->getActualHeight();
        return $this;
    }
    
    /**
     * Method responsible for creating document.
     * @return KrscReports_Document_Element_Default object on which method was executed
     */
    public function constructDocument() 
    {
        $this->_oBuilder->constructDocument();
        return $this;
    }
}
