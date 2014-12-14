<?php

/**
 * Table with possibility to add column sum at the end (extension from normal table builder). 
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableWithSums extends KrscReports_Builder_Excel_PHPExcel_ExampleTable implements KrscReports_Builder_Interface_Table
{
    /**
     * @var Array columns which are to sum 
     */
    protected $_aColumnsToSum;
    
    /**
     * @var Integer start height of rows (for formulas) 
     */
    protected $_iStartHeightOfRows;
    
    /**
     * 
     * @param String $sColumnName add column which will have sum formula under table rows
     */
    public function addColumnToSum( $sColumnName )
    {
        $this->_aColumnsToSum[] = $sColumnName;
    }
    
    public function setRows() 
    {
        $this->_iStartHeightOfRows = $this->_iActualHeight;
        parent::setRows();
    }
    
    public function endTable() 
    {
        parent::endTable();
        
        
        foreach( $this->_aColumnsToSum as $sColumnName )
        {   // add sum formula for each column
            // position of summed column
            $iPosition = array_search( $sColumnName, array_keys( $this->_aData[0] ) );
            
            $sColumnCoordinate = $this->_oCell->getColumnDimension( $iPosition );
            $sFormula = sprintf( '=SUM(%s%d:%s%d)', $sColumnCoordinate, $this->_iStartHeightOfRows, $sColumnCoordinate, ($this->_iActualHeight-1) );
            
            $this->_oCell->setValue( $sFormula );
            $this->_oCell->constructCell( $this->_iActualWidth +  $iPosition, $this->_iActualHeight );
        }
        
        $this->_iActualHeight++;
    }
}
    
