<?php
/**
 * Example report creating table in PHPExcel. 
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Report_ExampleReportVariousPlaces
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
        
        $oStyle = new KrscReports_Type_Excel_PHPExcel_Style();
        $oStyle->setStyleCollection( $oCollectionDefault );
        $oStyle->setStyleCollection( $oCollectionRow, KrscReports_Document_Element_Table::STYLE_ROW );
        
        //die( var_dump( $oStyle->getStyleArray( KrscReports_Document_Element_Table::STYLE_ROW ) ) );
        KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new PHPExcel() );
        $oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
        $oCell->setStyleObject( $oStyle );
        
        $oBuilder = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder->setCellObject( $oCell );
        $oBuilder->setData( array( array( 'Pierwsza kolumna' => '1', 'Druga kolumna' => '2' ), array( 'Pierwsza kolumna' => '3', 'Druga kolumna' => '4' ) ) );
        
        // creation of element responsible for creating table
        $oElementTable = new KrscReports_Document_Element_Table();
        $oElementTable->setBuilder( $oBuilder );
        $oElementTable->setStartWidth(1);
        
        $oBuilder2 = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder2->setCellObject( $oCell );
        $oBuilder2->setData( array( array( 'I kolumna' => '5', 'II kolumna' => '6' ), array( 'Pierwsza kolumna' => '7', 'Druga kolumna' => '8' ) ) );
        
        
        $oElementTable2 = new KrscReports_Document_Element_Table();
        $oElementTable2->setBuilder( $oBuilder2 );
        $oElementTable2->setStartWidth( 3 );
        $oElementTable2->setStartHeight( 15 );
        
        $oBuilder3 = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder3->setCellObject( $oCell );
        $oBuilder3->setData( array( array( 'I kolumna' => '55', 'II kolumna' => '66' ), array( 'Pierwsza kolumna' => '77', 'Druga kolumna' => '88' ) ) );
                
        $oElementTable3 = new KrscReports_Document_Element_Table();
        $oElementTable3->setBuilder( $oBuilder3 );
        $oElementTable3->setStartWidth( 6 );
        $oElementTable3->setStartHeight( 14 );
        
        
        $oBuilder4 = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder4->setCellObject( $oCell );
        $oBuilder4->setData( array( array( 'I kolumna' => '155', 'II kolumna' => '166' ), array( 'Pierwsza kolumna' => '177', 'Druga kolumna' => '188' ) ) );
                
        $oElementTable4 = new KrscReports_Document_Element_Table();
        $oElementTable4->setBuilder( $oBuilder4 );
        
        // adding table to spreadsheet
        $oElement = new KrscReports_Document_Element();
        $oElement->addElement( $oElementTable );
        $oElement->addElement( $oElementTable2 );
        $oElement->addElement( $oElementTable3 );
        $oElement->addElement( $oElementTable4 );
                
                
        $oElement->beforeConstructDocument();
        $oElement->constructDocument();
        $oElement->afterConstructDocument();
            
    }
}
