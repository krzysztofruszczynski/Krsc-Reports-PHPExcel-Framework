<?php
namespace KrscReports;

/**
 * Service for translating columns.
 */
class ColumnTranslatorService
{
    protected $translator;
    
    public function __construct($translator)
    {
        $this->translator = $translator;
    }
    
    /**
     * Method returning translator actually set in this service (no need to pass it seperately when this service is passed)
     * @return Object
     */
    public function getTranslator()
    {
        return $this->translator;
    }
    
    public function translateColumns($columns, $translatorDomain)
    {
        foreach ($columns as $key => $value) {
            $columns[$key] = $this->translator->trans($value, array(), $translatorDomain);
        }
        
        return $columns;
    }
}

