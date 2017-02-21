KrscReports
===========
In order to run, this project needs PHPExcel. Put PHPExcel source in Classes folder or install it via composer. Then load in your browser Examples.php file from the main folder and see the possibilities of this library. Live demo is here: [https://www.krsc.ruszczynski.eu/reports/example/list]

Install via composer
===========
To install this application via composer, create composer.json file with that content (or add it to existing composer.json file):
<pre>
{
"repositories": [
        {
            "type": "vcs",
            "url": "https://github.com/krzysztofruszczynski/Krsc-Reports-PHPExcel-Framework"
        }
    ],

"require-dev": {
        "krsc/krsc-reports-phpexcel-framework": "1.*"
    }
}
</pre>
Tested with Symfony 3.<br/>

The concept of application presented by tutorial
===========

Firstly, we have to set instance of PHPExcel class, which would be used inside called classes from KrscReports. We should do that by such command:<br/>
<code lang='php'>
  KrscReports_Builder_Excel_PHPExcel::setPHPExcelObject( new PHPExcel() );
</code>
<br/>Next step towards creating a worksheet is to create object, which will store all elements, that we want to display in generated file:<br/>
<code lang='php'>
  $oElement = new KrscReports_Document_Element();
</code>
<br/>In that moment, we can add elements to concrete worksheets. For example we want to add a table to sheet named "summary". Firstly, we should create object for that element:<br/>
<code lang='php'>
  $oElementTable = new KrscReports_Document_Element_Table();
</code>
<br/>At that moment we decided that we want to display a table, but we haven't decided yet precisiely, what would be the structure of the table (for example if we want a summary under last row for each column or we want first row before column name to use for table description). In order to do that, we have to create and set a builder like that:<br/>
<code lang='php'>
  $oBuilder = new KrscReports_Builder_Excel_PHPExcel_ExampleTable();
</code>
<br/>In this moment we have to add properties to builder. Firstly, we have to set an object, which would be responsible for creating cells:<br/>
<code lang='php'>
  $oCell = new KrscReports_Type_Excel_PHPExcel_Cell();
</code>
<br/>By calling this object, we can set styles for our builder, but using styles would be described later. Now we have to attach cell object to our builder:<br/>
<code lang='php'>
  $oBuilder->setCellObject( $oCell );
</code>
<br/>Next issue would be setting data for builder, for example like this:<br/>
<code lang='php'>
  $oBuilder->setData( array( array( 'First column' => '1', 'Second column' => '2' ),
                      array( 'First column' => '3', 'Second column' => '4' ) ) );
</code>
<br/>For creating a table, data structure should obey certain pattern. It should be an array, which subsequent elements are arrays (without fixed keys). Each subarray keys are names of columns and their values are what would be displayed in worksheet. Subarray keys are used to make header rows for tables.<br/>
Having properly configured builder we can attach it now to element:<br/>
<code lang='php'>
  $oElementTable->setBuilder( $oBuilder );
</code>
<br/>In that moment we have a proper element ready to be inserted in worksheet. Here we use element created at the beginning and add our table element to it:<br/>
<code lang='php'>
  $oElement->addElement( $oElementTable, 'first_one' );
</code>
<br/>Second parameter is the name of sheet, in which element would be displayed. If this parameter is omitted, default name for sheet would be used. It is possible to place many elements in one sheet. It is also possible to create many sheets. If user doesn't change it, the elements would be displayed one after the other.<br/>
After that, we have to construct document from our objects. In order to do that, we execute those 3 commands (order is important):<br/>
<code lang='php'>
  $oElement->beforeConstructDocument();
$oElement->constructDocument();
$oElement->afterConstructDocument();
</code>
<br/>In that moment we have correctly constructed PHPExcel object and we are ready to create a file from it. To do so, we have to create writer and save it, just like that:<br/>
<code lang='php'>
  $oWriter = PHPExcel_IOFactory::createWriter( KrscReports_Builder_Excel_PHPExcel::getPHPExcelObject(), 'Excel2007');
$oWriter->save('php://output');
</code>
<br/>That is all what we need to create a PHPExcel spreadsheet. Look for examples in library code if you want to extend your knowledge. Thanks for reading :)



