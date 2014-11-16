<?php
/**
 * Interface for creating tables.
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface KrscReports_Builder_Interface_Table {
    
    public function beginTable();
    
    public function setHeaderRow();
    
    public function setRows();
    
    public function endTable();
}
