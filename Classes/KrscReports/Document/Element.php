<?php
/**
 * Use of composite design pattern. 
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element
{
    /**
     * number of lines between elements in the spreadsheet
     */
    const LINES_BETWEEN_ELEMENTS = 2;
    
    /**
     * default name of group
     */
    const DEFAULT_GROUP_NAME = 'document1';
    
    /**
     * @var Array array where group name is a key and value is an array with elements
     */
    protected $_aElements = array();
    
    /**
     * @var Array array where group name is a key and value is actual width for this group
     */
    protected $_aActualWidths = array();
    
    /**
     * @var Array array where group name is a key and value is actual height for this group
     */
    protected $_aActualHeights = array();
    
    /**
     * @var String name of group to which belongs this element 
     */
    protected $_sGroupName = self::DEFAULT_GROUP_NAME;
    /* później
    public function addOrModifyGroup( $sGroupName, $iNumber )
    {
        
    }
    */
    public function addElement( $oElement, $sGroupName = self::DEFAULT_GROUP_NAME )
    {
        $this->_aElements[$sGroupName][] = $oElement;
        return $this;
    }
    
    public function setGroupName( $sGroupName )
    {
        $this->_sGroupName = $sGroupName;
        return $this;
    }
    
    public function afterConstructDocument() { return $this; }
    
    public function beforeConstructDocument() { return $this; }
    
    /**
     * Method iterating over all elements in composite
     */
    public function constructDocument()
    {
        // iteration over groups
        foreach( $this->_aElements as $sGroupName => $aGroupElements )
        {   
            // iteration over elements in group
            foreach( $aGroupElements as $oElement )
            {
                $oElement->setInnerGroupName( $sGroupName );
                
                // setting place of an element
                $oElement->setStartHeight( isset( $this->_aActualHeights[$sGroupName] ) ? $this->_aActualHeights[$sGroupName] + self::LINES_BETWEEN_ELEMENTS : 1, false );
                
                $oElement->beforeConstructDocument();
                $oElement->constructDocument();
                $oElement->afterConstructDocument();
                
                $this->_aActualHeights[$sGroupName] = $oElement->getActualHeight();
            }
        }
        
        return $this;
    }
}
?>
