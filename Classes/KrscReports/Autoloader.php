<?php
KrscReports_Autoloader::Register();

/**
 * Autoloader for KrscReports (based on PHPExcel_Autoloader)
 *
 * @category KrscReports
 * @package KrscReports
 * @copyright   Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class KrscReports_Autoloader
{
    /**
     * Register the Autoloader with SPL
     *
     */
    public static function Register() {
        if (function_exists('__autoload')) {
            //    Register any existing autoloader function with SPL, so we don't get any clashes
            spl_autoload_register('__autoload');
        }
        //    Register ourselves with SPL
        return spl_autoload_register(array('KrscReports_Autoloader', 'Load'));
    }   //    function Register()


    /**
     * Autoload a class identified by name
     *
     * @param    string    $pClassName        Name of the object to load
     */
    public static function Load($pClassName){
        if ((class_exists($pClassName,FALSE)) || (strpos($pClassName, 'KrscReports') !== 0)) {
            //    Either already loaded, or not a PHPExcel class request
            return FALSE;
        }

        $pClassFilePath = PHPEXCEL_ROOT .
                          str_replace('_',DIRECTORY_SEPARATOR,$pClassName) .
                          '.php';

        if ((file_exists($pClassFilePath) === FALSE) || (is_readable($pClassFilePath) === FALSE)) {
            //    Can't load
            return FALSE;
        }

        require($pClassFilePath);
    }   //    function Load()

}
