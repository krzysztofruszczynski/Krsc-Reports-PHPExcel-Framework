<?php
namespace KrscReports\Type\Excel\PHPExcel\Style\Bundle;

abstract class AbstractStyle
{
    /**
     * @var \KrscReports_Type_Excel_PHPExcel_Style 
     */
    protected $style;
    
    abstract protected function setCollections();
    
    /**
     * To be override in non abstract classes.
     * @param \KrscReports_Type_Excel_PHPExcel_Style 
     */
    abstract protected function setStyleObject( $style );
    
    /**
     * 
     * @return \KrscReports_Type_Excel_PHPExcel_Style
     */
    public function getStyleObject()
    {
        if ( !isset( $this->collectionDefault ) ) {
            $this->setCollections();
        }
        
        if ( !isset( $this->style ) ) {
            $this->style = new \KrscReports_Type_Excel_PHPExcel_Style();
            $this->setStyleObject( $this->style );
        }
        
        return $this->style;
    }
}
