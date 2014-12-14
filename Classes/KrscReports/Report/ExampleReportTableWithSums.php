<?php
/**
 * Example report creating table in PHPExcel. 
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Report_ExampleReportTableWithSums
{
    /**
     * @return Void
     */
    public function generate()
    {
        // setting styles - adding elements to iterator (for this moment styles are global)
        $oCollectionDefault = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionDefault->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders() );
        
        $oCollectionRow = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionRow->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill() );
        $oCollectionRow->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        $oFill = new KrscReports_Type_Excel_PHPExcel_Style_Fill_ExampleFill();
        $oFill->setFillColor( 'CCCCCC' ); // gray
        $oCollectionSummary = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollectionSummary->addStyleElement( $oFill );
        $oCollectionSummary->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        
        
        $oStyle = new KrscReports_Type_Excel_PHPExcel_Style();
        $oStyle->setStyleCollection( $oCollectionDefault );
        $oStyle->setStyleCollection( $oCollectionRow, KrscReports_Document_Element_Table::STYLE_ROW );
        $oStyle->setStyleCollection( $oCollectionSummary, KrscReports_Document_Element_Table::STYLE_SUMMARY );
        
        //die( var_dump( $oStyle->getStyleArray( KrscReports_Document_Element_Table::STYLE_ROW ) ) );
        KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new PHPExcel() );
        $oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
        $oCell->setStyleObject( $oStyle );
        
        $oBuilder = new KrscReports_Builder_Excel_PHPExcel_TableWithSums();
        $oBuilder->setCellObject( $oCell );
        $oBuilder->setData( array( array( 'Pierwsza kolumna' => '1', 'Druga kolumna' => '2' ), array( 'Pierwsza kolumna' => '3', 'Druga kolumna' => '4' ) ) );
        $oBuilder->addColumnToSum( 'Druga kolumna' );
        
        // creation of element responsible for creating table
        $oElementTable = new KrscReports_Document_Element_Table();
        $oElementTable->setBuilder( $oBuilder );
        
        $oBuilder2 = new KrscReports_Builder_Excel_PHPExcel_TableWithSums();
        $oBuilder2->setCellObject( $oCell );
        $oBuilder2->setData( array( array( 'I kolumna' => '5', 'II kolumna' => '6' ), array( 'I kolumna' => '7', 'II kolumna' => '8' ), array( 'I kolumna' => '-1', 'II kolumna' => '4' ) ) );
        $oBuilder2->addColumnToSum( 'I kolumna' );
        $oBuilder2->addColumnToSum( 'II kolumna' );
        
        $oElementTable2 = new KrscReports_Document_Element_Table();
        $oElementTable2->setBuilder( $oBuilder2 );
        
        $oBuilder3 = new KrscReports_Builder_Excel_PHPExcel_TableWithSums();
        $oBuilder3->setCellObject( $oCell );
        $oBuilder3->setData( array( array( 'I kolumna' => '5', 'II kolumna' => '6' ), array( 'I kolumna' => '7', 'II kolumna' => '8' ) ) );
        $oBuilder3->addColumnToSum( 'I kolumna' );
        
        $oElementTable3 = new KrscReports_Document_Element_Table();
        $oElementTable3->setBuilder( $oBuilder3 );
        
        // adding table to spreadsheet
        $oElement = new KrscReports_Document_Element();
        $oElement->addElement( $oElementTable );
        $oElement->addElement( $oElementTable2 );
        $oElement->addElement( $oElementTable3 );
        
                
        $oElement->beforeConstructDocument();
        $oElement->constructDocument();
        $oElement->afterConstructDocument();
            
    }
}

