<?php

namespace tricktrick\phpexcel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Class PhpExcel
 * @property Drawing $objDrawing
 */
class PhpExcel extends \yii\base\BaseObject
{
    /**
     * @var string
     */
    public $defaultFormat = 'Excel2007';

    /**
     * Creates new workbook
     * @return Spreadsheet
     */
    public function create()
    {
        return new Spreadsheet();
    }

	/**
	 * Creates new Worksheet Drawing
	 * @return Drawing
	 */
	public function getObjDrawing() {
		return new Drawing();
	}

    /**
     * @param $filename
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function load($filename)
    {
        return IOFactory::load($filename);
    }

    /**
     * @param Spreadsheet $object
     * @param $filename
     * @param null $format
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws \yii\base\ExitException
     * @throws \yii\web\RangeNotSatisfiableHttpException
     */
    public function responseFile(Spreadsheet $object, $filename, $format = null)
    {
        if ($format === null) {
            $format = $this->resolveFormat($filename);
        }
        $writer = IOFactory::createWriter($object, $format);
        ob_start();
        $writer->save('php://output');
        $content = ob_get_clean();
        \Yii::$app->response->sendContentAsFile($content, $filename, $this->resolveMime($format));
        \Yii::$app->end();
    }

    /**
     * @param $sheet
     * @param $data
     * @param $config
     * @return mixed
     */
    public function writeSheetData($sheet, $data, $config)
    {
        $config['sheet'] = &$sheet;
        $config['data'] = $data;
        $writer = new ExcelDataWriter($config);
        $writer->write();
        return $sheet;
    }

    /**
     *
     */
    public function writeTemplateData(/* TODO */)
    {
        // TODO: implement
    }

    /**
     * @param $sheet
     * @param $config
     */
    public function readSheetData($sheet, $config)
    {
        // TODO: implement
    }

    /**
     *
     * @param $format
     * @return string
     */
    protected function resolveMime($format)
    {
        $list = [
            'CSV' => 'text/csv',
            'HTML' => 'text/html',
            'PDF' => 'application/pdf',
            'OpenDocument' => 'application/vnd.oasis.opendocument.spreadsheet',
            'Excel5' => 'application/vnd.ms-excel',
            'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];
        return isset($list[$format]) ? $list[$format] : 'application/octet-stream';
    }

    /**
     *
     * @param $filename
     * @return string
     */
    protected function resolveFormat($filename)
    {
        // see IOFactory::createReaderForFile etc.
        $list = [
            'ods' => 'OpenDocument',
            'xls' => 'Excel5',
            'xlsx' => 'Excel2007',
            'csv' => 'CSV',
            'pdf' => 'PDF',
            'html' => 'HTML',
        ];
        // TODO: check strtolower
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        return isset($list[$extension]) ? $list[$extension] : $this->defaultFormat;
    }
}
