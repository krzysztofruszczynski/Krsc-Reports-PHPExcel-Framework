<?php

/**
 * Table with possibility to add column sum at the end (extension from normal table builder). 
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Builder_Excel_PHPExcel_TableWithSums extends KrscReports_Builder_Excel_PHPExcel_ExampleTable implements KrscReports_Builder_Interface_Table
{
    /**
     * @var array columns which are to sum (each column name is next array value
     */
    protected $_aColumnsToSum;
    
    /**
     * @var integer height of first row with data (to be used by formulas) 
     */
    protected $_iStartHeightOfRows;
    
    /**
     * Method for adding columns, which will have sum formula under rows.
     * @param string $sColumnName name of the column which will have sum formula under table rows
     * @return KrscReports_Builder_Excel_PHPExcel_TableWithSums object on which method was executed
     */
    public function addColumnToSum( $sColumnName )
    {
        $this->_aColumnsToSum[] = $sColumnName;
        return $this;
    }
    
    /**
     * Action done when table rows has to be displayed (increments height with each row, remembers height of first row with data).
     * @return void
     */
    public function setRows() 
    {
        $this->_iStartHeightOfRows = $this->_iActualHeight;
        parent::setRows();
    }
    
    /**
     * Action done when table ends. If columns for sum formula have been added, it creates a field under the table with sum formula
     * (under particular column).
     * @return void
     */
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
    
