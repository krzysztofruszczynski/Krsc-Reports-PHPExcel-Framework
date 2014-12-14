<?php
/**
 * Interface for builders responsible for creating tables.
 * 
 * @category KrscReports
 * @package KrscReports_Builder
 * @author Krzysztof Ruszczyński <http://www.ruszczynski.eu>
 */
interface KrscReports_Builder_Interface_Table {
    
    /**
     * Action done when table begins.
     * @return void
     */
    public function beginTable();
    
    /**
     * Action done when table header row with name of the columns has to be displayed.
     * @return void
     */
    public function setHeaderRow();
    
    /**
     * Action done when table rows has to be displayed.
     * @return void
     */
    public function setRows();
    
    /**
     * Action done when table ends.
     * @return void
     */
    public function endTable();
}
