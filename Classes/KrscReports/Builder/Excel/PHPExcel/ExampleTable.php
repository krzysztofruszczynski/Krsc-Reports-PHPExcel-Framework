<?php

/**
  @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_ExampleTable extends KrscReports_Builder_Excel_PHPExcel implements KrscReports_Builder_Interface_Table
{
    
    public function beginTable() 
    {
        
    }

    public function endTable() 
    {
        
    }

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
    
    public function constructDocument()    
    {
        
    }

}
