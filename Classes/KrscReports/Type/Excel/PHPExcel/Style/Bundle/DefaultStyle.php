<?php
namespace KrscReports\Type\Excel\PHPExcel\Style\Bundle;

class DefaultStyle extends AbstractStyle
{    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionDefault;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionGray;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionBold;
    
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection if collection is shared, can be set in Collection folder
     */
    protected $collectionLightGrayBorders;
    
    /**
     * Method for setting collections (loaded from external files are created here).
     */
    protected function setCollections()
    {
        $this->collectionDefault = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionDefault->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_ExampleBorders() );

        $fillGray = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $fillGray->setColor( \KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_GRAY );

        $this->collectionGray = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionGray->addStyleElement( $fillGray );
        $this->collectionGray->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DoubleBorders() );
        $this->collectionGray->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

        $this->collectionBold = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionBold->addStyleElement( $fillGray );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DoubleBorders() );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Font_Bold() );
        $this->collectionBold->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

        $fillLightGray = new \KrscReports_Type_Excel_PHPExcel_Style_Fill_Basic();
        $fillLightGray->setColor( \KrscReports_Type_Excel_PHPExcel_Style_Default::COLOR_LIGHT_GRAY );
        $this->collectionLightGrayBorders = new \KrscReports_Type_Excel_PHPExcel_Style_Iterator_Collection();
        $this->collectionLightGrayBorders->addStyleElement( $fillLightGray );
        $this->collectionLightGrayBorders->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Borders_DashDotDotBorders() );
        $this->collectionLightGrayBorders->addStyleElement( new \KrscReports_Type_Excel_PHPExcel_Style_Alignment_Wrap() );

    }
    
    /**
     * 
     * @param \KrscReports_Type_Excel_PHPExcel_Style $style
     */
    protected function setStyleObject( $style )
    {
        $style->setStyleCollection( $this->collectionBold );
        $style->setStyleCollection( $this->collectionLightGrayBorders, \KrscReports_Document_Element_Table::STYLE_ROW );
        $style->setStyleCollection( $this->collectionGray, \KrscReports_Document_Element_Table::STYLE_HEADER_ROW );
    }

}