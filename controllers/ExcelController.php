<?php

namespace app\controllers;

use Yii;
use yii\web\Controller;
use yii\web\UploadedFile;
use app\models\UploadForm;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class ExcelController extends Controller
{
    public function actionUpload()
    {
        $model = new UploadForm();

        if (Yii::$app->request->isPost) {
            $model->file = UploadedFile::getInstance($model, 'file');
            if ($model->validate()) {
                // $filePath = 'uploads/' . $model->file->baseName . '.' . $model->file->extension;
                // $model->file->saveAs($filePath);
                $filePath = Yii::getAlias('@webroot/uploads/') . $model->file->baseName . '.' . $model->file->extension;
                $model->file->saveAs($filePath);

                // อ่านข้อมูลจากไฟล์ Excel
                $spreadsheet = IOFactory::load($filePath);
                $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

                // บันทึกข้อมูลลงฐานข้อมูล
                foreach ($sheetData as $row) {
                    // สมมติว่าฐานข้อมูลมีคอลัมน์ 'name', 'email', 'phone'
                    $command = Yii::$app->db->createCommand()->insert('data_excel', [
                        'name' => $row['A'],
                        'email' => $row['B'],
                        'phone' => $row['C'],
                    ]);
                    $command->execute();
                }

                Yii::$app->session->setFlash('success', 'Data uploaded successfully!');
                return $this->redirect(['upload']);
            }
        }

        return $this->render('upload', ['model' => $model]);
    }

    public function actionDownload()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        $data = Yii::$app->db->createCommand('SELECT * FROM data_excel')->queryAll();

        // สร้าง Spreadsheet
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // ตั้งชื่อหัวตาราง
        $sheet->setCellValue('A1', 'ID');
        $sheet->setCellValue('B1', 'Name');
        $sheet->setCellValue('C1', 'Email');
        $sheet->setCellValue('D1', 'Phone');

        // เพิ่มข้อมูลลงในไฟล์
        $rowNumber = 2; // เริ่มจากแถวที่ 2
        foreach ($data as $row) {
            $sheet->setCellValue('A' . $rowNumber, $row['id']);
            $sheet->setCellValue('B' . $rowNumber, $row['name']);
            $sheet->setCellValue('C' . $rowNumber, $row['email']);
            $sheet->setCellValue('D' . $rowNumber, $row['phone']);
            $rowNumber++;
        }

        // ตั้งชื่อไฟล์
        $fileName = 'data_excel_' . date('Y-m-d_H-i-s') . '.xlsx';

        // เขียนไฟล์และดาวน์โหลด
        $writer = new Xlsx($spreadsheet);
        $filePath = Yii::getAlias('@webroot/uploads/') . $fileName;
        $writer->save($filePath);

        return Yii::$app->response->sendFile($filePath)->on(\yii\web\Response::EVENT_AFTER_SEND, function ($event) {
            unlink($event->data);
        }, $filePath);
    }
}
