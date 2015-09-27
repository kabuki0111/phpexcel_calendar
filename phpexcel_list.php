<?php
require_once 'PHPExcel.php';
require_once 'PHPExcel/IOFactory.php';


//ブラウザへ出力をリダイレクト
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="colender.xlsx"');
header('Cache-Control: max-age=0');


define("UNIX_HOUR", 3600);


//オブジェクトの生成
$xl = new PHPExcel();

$trg_year = 2015;
$date_arr = array();

//一年間の日付を計算し配列に代入(１月から１２月まで）
for($i_mouth=1;$i_mouth<=12;$i_mouth++){
	$xl->setActiveSheetIndex($i_mouth-1);
	$sheet = $xl->getActiveSheet();
	$sheet->setTitle("スケジュール_{$i_mouth}月");
	
	$trg_day = 1;
	while(1){
		if(!checkdate($i_mouth, $trg_day, $trg_year)){
			break;
		}

		$date_arr[] = array(
				"year" => $trg_year,
				"mouth" => $i_mouth,
				"day" => $trg_day
		);
		
		$trg_day++;
	}
	
	$xl->createSheet();
}


//指定の年月の一覧を表示
$index_pos_x = 2;
$trg_mouth = 1;
foreach ($date_arr as $key=>$value){
	if($value["mouth"] != $trg_mouth){
		$index_pos_x = 2;
		$trg_mouth = $value["mouth"];
	}
	
	$xl->setActiveSheetIndex($value["mouth"]-1);
	$sheet = $xl->getActiveSheet();
	
	//各日付を一時間ごとに設定
	$strdate = "{$value["mouth"]}/{$value["day"]}";
	$sheet->setCellValueByColumnAndRow($index_pos_x, 2, $strdate);
	$sheet->setCellValueByColumnAndRow($index_pos_x+1, 2, "予定");
	
	for($i_hour=0;$i_hour<24;$i_hour++){
		$set_date = date("{$value["year"]}/{$value["mouth"]}/{$value["day"]} {$i_hour}:00:00");
		$trg_unix = strtotime($set_date);
		$trg_date = date("H", $trg_unix);
		$sheet->setCellValueByColumnAndRow($index_pos_x, $i_hour+3, intval($trg_date)."時");
	}
	$index_pos_x += 2;
}


//Excel2007形式で保存
$writer = PHPExcel_IOFactory::createWriter($xl, 'Excel2007');
$writer->save('php://output');
exit;
?>