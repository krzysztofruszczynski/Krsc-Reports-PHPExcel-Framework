<?php
namespace KrscReports\Views;

/**
 * Report consisting of table and another table with filters used for the first one.
 */
class SingleTableWithFilters extends SingleTable
{
    /**
     * compatible with value 
     */
    const KEY_FILTERS = 'filters';
    
    /**
     * key in options for array with selected filters values
     */
    const KEY_FILTERS_VALUES = 'filters_values';
    
    /**
     * key in options for filter table title
     */
    const KEY_FILTER_TABLE_TITLE = 'title_filter_table';
    
    protected function getFilterTableBuilder()
    {
        $translatedColumns = $this->columnTranslator->translateColumns( $this->options[self::KEY_FILTERS], $this->options[self::KEY_TRANSLATOR_DOMAIN] );
        $legendRowFilterValues = $this->options[self::KEY_FILTERS_VALUES];
        
        $filterData = $this->addColumnNames( array( $legendRowFilterValues ), $translatedColumns );
        
        $builder = new \KrscReports_Builder_Excel_PHPExcel_TableTitle();
        $builder->setTitle( ( isset( $this->options[self::KEY_SINGLE_TABLE_TITLE] ) ? $this->options[self::KEY_SINGLE_TABLE_TITLE] : 'Filters' ) );
        $builder->setCellObject( $this->getCell() );        
        $builder->setData( $filterData );
        
        return $builder;
    }
    
    public function getDocumentElement( $spreadsheetName = 'Results with filters' )
    {
        $builderLegend = $this->getFilterTableBuilder();
        $builder = $this->getSingleTableBuilder( $spreadsheetName );
        
        $this->options[self::KEY_COLUMN_LINES_BETWEEN_ELEMENTS] = 1;
        
        $this->setTableElement($builderLegend, $spreadsheetName);
        $this->setTableElement($builder, $spreadsheetName);
        
        return $this->documentElement; 
    }
}

