<?php
use yii\widgets\ActiveForm;
use yii\helpers\Html;

$this->title = 'Upload Excel File';
?>
<h1><?= Html::encode($this->title) ?></h1>

<?php $form = ActiveForm::begin([
    'options' => ['enctype' => 'multipart/form-data']
]); ?>

<?= $form->field($model, 'file')->fileInput() ?>

<div class="form-group">
    <?= Html::submitButton('Upload', ['class' => 'btn btn-primary']) ?>
</div>

<div class="form-group">
    <?= Html::a('Download Excel', ['excel/download'], ['class' => 'btn btn-success']) ?>
</div>


<?php ActiveForm::end(); ?>
