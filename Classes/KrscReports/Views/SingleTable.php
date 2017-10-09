<?php
namespace KrscReports\Views;

/**
 * Report consisting of one single table.
 */
class SingleTable extends AbstractView
{
    /**
     * key style with gray fill
     */
    const STYLE_CELL_GRAY = 'style_gray';

    /**
     * key style for light gray fill with double borders
     */
    const STYLE_CELL_LIGHT_GRAY = 'style_light_gray';

    /**
     * key in options for single table title
     */
    const KEY_SINGLE_TABLE_TITLE = 'title_single_table';

    /**
     * Method setting table builder in document element, setting necessary options (like lines between elements) 
     * and adding it to internal variable later used by Krsc Reports.
     * @param \KrscReports_Builder_Excel_PHPExcel_TableCurrent $builder previously created and configured table builder
     * @param string $spreadsheetName spreadsheet name, where element would be placed
     */
    protected function setTableElement($builder, $spreadsheetName)
    {
        $elementTable = new \KrscReports_Document_Element_Table();
        $elementTable->setBuilder($builder);

        if ( isset($this->options[self::KEY_COLUMN_LINES_BETWEEN_ELEMENTS]) ) {
            $elementTable->setLinesBetweenElements($this->options[self::KEY_COLUMN_LINES_BETWEEN_ELEMENTS]);
        }
        
        // adding table to spreadsheet
        $this->documentElement->addElement($elementTable, $spreadsheetName);
    }
    
    /**
     * Builder for single table.
     * @param string $spreadsheetName spreadsheet name (for setting column widths in proper spreadsheet)
     * @return \KrscReports_Builder_Excel_PHPExcel_TableCurrent builder for report
     */
    protected function getSingleTableBuilder( $spreadsheetName )
    {
        $builder = new \KrscReports_Builder_Excel_PHPExcel_TableCurrent();
        $builder->setAutoFilter( ( isset( $this->options[self::KEY_AUTOFILTER] ) ? $this->options[self::KEY_AUTOFILTER] : true ) );

        $builder->setGroupName( $spreadsheetName ); // change active sheet
        $builder->setCellObject( $this->getCell() );

        if ( isset( $this->options[self::KEY_SINGLE_TABLE_TITLE] ) ) {
            $builder->setTitle( $this->options[self::KEY_SINGLE_TABLE_TITLE] );
        }

        $builder->setData( $this->data );
        $builder->setColumnNames( $this->columnNames );

        return $builder;
    }

    /**
     * Method generating report.
     */
    public function getDocumentElement( $spreadsheetName = 'Results' )
    {
        $builder = $this->getSingleTableBuilder( $spreadsheetName );
        $this->setTableElement($builder, $spreadsheetName);

        return $this->documentElement;
    }

    public function getDescription()
    {
        return 'Single table view.';
    }
}
