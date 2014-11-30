<?php
/**
 * Example report creating table in PHPExcel. 
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Report_ExampleReport
{
    /**
     * @return Void
     */
    public function generate()
    {
        // setting styles - adding elements to iterator (for this moment styles are global)
        $oCollection = new KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $oCollection->addStyleElement( new KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders() );
        
        $oStyle = new KrscReports_Type_Excel_PHPExcel_Style();
        $oStyle->setStyleCollection( $oCollection );
        
        KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new PHPExcel() );
        $oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
        $oCell->setStyleObject( $oStyle );
        
        $oBuilder = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder->setCellObject( $oCell );
        $oBuilder->setData( array( array( 'Pierwsza kolumna' => '1', 'Druga kolumna' => '2' ), array( 'Pierwsza kolumna' => '3', 'Druga kolumna' => '4' ) ) );
        
        // creation of element responsible for creating table
        $oElementTable = new KrscReports_Document_Element_Table();
        $oElementTable->setBuilder( $oBuilder );
        
        
        $oBuilder2 = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
        $oBuilder2->setCellObject( $oCell );
        $oBuilder2->setData( array( array( 'Pierwsza kolumna' => '5', 'Druga kolumna' => '6' ), array( 'Pierwsza kolumna' => '7', 'Druga kolumna' => '8' ) ) );
        
        
        $oElementTable2 = new KrscReports_Document_Element_Table();
        $oElementTable2->setBuilder( $oBuilder2 );
       
        // adding table to spreadsheet
        $oElement = new KrscReports_Document_Element();
        $oElement->addElement( $oElementTable );
        $oElement->addElement( $oElementTable2, 'Second_one' );
        
        
                
        $oElement->beforeConstructDocument();
        $oElement->constructDocument();
        $oElement->afterConstructDocument();
            
    }
}

