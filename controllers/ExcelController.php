<?php

namespace app\controllers;

use Yii;
use yii\web\Controller;
use yii\web\UploadedFile;
use app\models\UploadForm;
use PhpOffice\PhpSpreadsheet\IOFactory;

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
}
