<?php

/**
 * @todo: może tworzenie samej tabeli wrzucić gdzieś dalej
 * Use of builder design pattern.
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
class KrscReports_Document_Element_Table extends KrscReports_Document_Element_Default
{
    
    public function constructDocument() 
    {
        $this->_oBuilder->setGroupName( $this->_sGroupName );
        $this->_oBuilder->beginTable();
        $this->_oBuilder->setHeaderRow();
        $this->_oBuilder->setRows();
        $this->_oBuilder->endTable();
    }
}

