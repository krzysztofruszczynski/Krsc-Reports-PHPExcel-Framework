<?php

/**
 * Use of builder design pattern.
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element_Table extends KrscReports_Document_Element_Default
{
    /**
     * style name for beginning of table
     */
    const STYLE_BEGIN_TABLE = 'beginTable';
    
    /**
     * style name for header with rows
     */
    const STYLE_HEADER_ROW = 'headerRow';
    
    /**
     * style name for normal row
     */
    const STYLE_ROW = 'row';
    
    /**
     * style name for summary row (for example with sums)
     */
    const STYLE_SUMMARY = 'summary';
    
    
    public function constructDocument() 
    {
        $this->_oBuilder->setStyleKey( KrscReports_Type_Excel_PHPExcel_Style::KEY_DEFAULT );
        
        $this->_oBuilder->setGroupName( $this->_sGroupName );
        $this->_oBuilder->beginTable();
        
        $this->_oBuilder->setStyleKey( self::STYLE_HEADER_ROW );
        $this->_oBuilder->setHeaderRow();
        
        $this->_oBuilder->setStyleKey( self::STYLE_ROW );
        $this->_oBuilder->setRows();
        
        $this->_oBuilder->setStyleKey( self::STYLE_SUMMARY );
        $this->_oBuilder->endTable();
    }
}

