<?php

/**
 * Builder responsible for creating example table.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_ExampleTable extends KrscReports_Builder_Excel_PHPExcel implements KrscReports_Builder_Interface_Table
{
    /**
     * Action done when table begins.
     * @return void
     */
    public function beginTable() 
    {
        
    }

    /**
     * Action done when table ends.
     * @return void
     */
    public function endTable() 
    {
        
    }

    /**
     * Action done when table header row with name of the columns has to be displayed (increments height).
     * @return void
     */
    public function setHeaderRow() 
    {
        $iIterator = 0;
        
        foreach( $this->_aData[0] as $sColumnName => $mColumnValue )
        {
            $this->_oCell->setValue( $sColumnName );
            $this->_oCell->constructCell( $this->_iActualWidth + $iIterator++, $this->_iActualHeight );
        }
        
        // adding one row in registry
        $this->_iActualHeight++;
    }

    /**
     * Action done when table rows has to be displayed (increments height with each row).
     * @return void
     */
    public function setRows() 
    {
        
        foreach( $this->_aData as $aRow )
        {   // iterating over rows
            $iIterator = 0;
            foreach( $aRow as $mColumnValue )
            {
                $this->_oCell->setValue( $mColumnValue );
                $this->_oCell->constructCell( $this->_iActualWidth + $iIterator++, $this->_iActualHeight );
            }
            
            // adding one row in registry
            $this->_iActualHeight++;
        }
    }
    
    /**
     * Action responsible for creating document (while creating table not used but has to be implemented).
     * @return void
     */
    public function constructDocument()    
    {
        
    }

}
