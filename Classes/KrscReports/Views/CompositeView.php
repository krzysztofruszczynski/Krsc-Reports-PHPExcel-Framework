<?php
namespace KrscReports\Views;

/**
 * Report consisting of multiple views inheriting from AbstractView.
 */
class CompositeView extends AbstractView
{
    /**
     * @var array Views hierarchy (input from user):
     * - spreadsheet name
     *      - instances of abstract view 
     */
    protected $viewArray = array();
    
    /**
     * Method adding element to composer.
     * @param \KrscReports\Views\AbstractView $view view to add
     * @param string $spreadsheetName name of spreadsheet where data should be added
     * @return CompositeView object on which this method is executed
     */
    public function addView( AbstractView $view, $spreadsheetName = 'composite' )
    {
        $this->viewArray[$spreadsheetName][] = $view;
        return $this;
    }
    
    public function getDocumentElement( $spreadsheetName = 'worksheet' )
    {
        foreach ( $this->viewArray as $spreadsheetName => $viewCollection ) {            
            foreach ( $viewCollection as $view ) {   
                $viewOptions = $view->getOptions();
                $view->addOptions( $this->options ); // overrides previously set options
                $view->addOptions( $viewOptions ); // prevent from overriding previously set options
                
                $view->setDocumentElement( $this->documentElement ); // there is a reference so object is automatically changed
                $view->getDocumentElement( $spreadsheetName );
            }
        }
        
        return $this->documentElement;
    }
}
