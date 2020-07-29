<?php

namespace app\controller;

use AipOcr;
use app\BaseController;
use Imagick;
use think\Exception;
use think\facade\Log;
use think\facade\View;
use think\Request;

class Index extends BaseController
{
	
	//上传图片页面
	public function upload()
	{
		View::assign('data', 'ThinkPHP');
		// 模板输出
		return View::fetch('index');
	}
	public function mult_scan(Request $request){
		$files = $request->file('images');
		foreach($files as $file){
			$this->change($file);
		}
	}
	//将pdf转换为单张图片
	public function change($file){
		//源文件pdf名称
		$originalname = $file->getOriginalName();
		//上传文件
		$savename = \think\facade\Filesystem::putFileAs($originalname, $file, $originalname);
		$dir      = app()->getRuntimePath() . 'storage/';
		$filepath = $dir . $savename;
		try {
			if (!extension_loaded('imagick')) {
				return false;
			}
			$im = new Imagick();
			$im->setResolution(300, 300); //设置分辨率 值越大分辨率越高
			$im->setCompressionQuality(40);
			$im->readImage($filepath);
			foreach ($im as $k => $v) {
				$v->setImageFormat('png');
				$fileName = $dir . $originalname . '/' . $originalname . '(' . $v->getIteratorIndex() . ').png';
				if (file_exists($fileName)) {
					$return[] = $fileName;
					continue;
				}
				if ($v->writeImage($fileName) == true) {
					$return[] = $fileName;
				}
			}
			$txtname = $dir . $originalname . '/' . $originalname . '.txt';
			if (file_exists($txtname)) {
				$result = file_get_contents($txtname);
			} else {
				//调用ocr
				$result = $this->getString($return);
				//将获取的字符串保存
				file_put_contents($dir . $originalname . '/' . $originalname . '.txt', $result);
			}
			//提取字段
			$result = $this->getTieTong($result);
			//填入execl
			$this->inputexecl($result);
		} catch (Exception $e) {
			Log::error($originalname.$e->getMessage().$e->getLine());
		}
	}
	//将pdf转换为单张图片
	public function scan(Request $request)
	{
		$file = $request->file('image');
		//源文件pdf名称
		$originalname = $file->getOriginalName();
		//上传文件
		$savename = \think\facade\Filesystem::putFileAs($originalname, $file, $originalname);
		$dir      = app()->getRuntimePath() . 'storage/';
		$filepath = $dir . $savename;
		try {
			if (!extension_loaded('imagick')) {
				return false;
			}
			$im = new Imagick();
			$im->setResolution(300, 300); //设置分辨率 值越大分辨率越高
			$im->setCompressionQuality(40);
			$im->readImage($filepath);
			foreach ($im as $k => $v) {
				$v->setImageFormat('png');
				$fileName = $dir . $originalname . '/' . $originalname . '(' . $v->getIteratorIndex() . ').png';
				if (file_exists($fileName)) {
					$return[] = $fileName;
					continue;
				}
				if ($v->writeImage($fileName) == true) {
					$return[] = $fileName;
				}
			}
			$txtname = $dir . $originalname . '/' . $originalname . '.txt';
			if (file_exists($txtname)) {
				$result = file_get_contents($txtname);
			} else {
				//调用ocr
				$result = $this->getString($return);
				//将获取的字符串保存
				file_put_contents($dir . $originalname . '/' . $originalname . '.txt', $result);
			}
			//提取字段
			$result = $this->getTieTong($result);
			//填入execl
			$this->inputexecl($result);
		} catch (Exception $e) {
			dump($e);
			Log::error($originalname.$e->getMessage().'line:'.$e->getLine());
		}
		dump('转换成功');
		return view('index', [
			'data' => $result,
		]);
	}
	//数据插入execl
	public function inputexecl($result)
	{
		$reader             = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
		$spreadsheet        = $reader->load('/Users/kingsley/PhpstormProjects/scanpdf/public/铁通合同100份-合同信息截取测试.xlsx'); //载入excel表格
		$worksheet          = $spreadsheet->getActiveSheet();
		$highestRow         = $worksheet->getHighestRow(); // 总行数
		$highestColumn      = $worksheet->getHighestColumn(); // 总列数
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
		$lines              = $highestRow - 2;
		if ($lines <= 0) {
			exit('Excel表格中没有数据');
		}
		for ($row = 2; $row <= $highestRow; ++$row) {
			$num=$worksheet->getCellByColumnAndRow(3, $row)->getValue();
			$num=explode("-",$num);
			if ($num[2] ==  $result['合同编号']) {
				//已扫描情况
				$worksheet->setCellValueByColumnAndRow(5, $row, '已检测');
				//地址
				$worksheet->setCellValueByColumnAndRow(6, $row, $result['地址']);
				//租赁期限
				$worksheet->setCellValueByColumnAndRow(7, $row, $result['租赁期限']);
				//场地开始使用时间
				$worksheet->setCellValueByColumnAndRow(8, $row, $result['场地开始使用时间']);
				//交租金日
				$worksheet->setCellValueByColumnAndRow(9, $row, $result['交租金日']);
				//租金金额
				$worksheet->setCellValueByColumnAndRow(10, $row, $result['租金']);
				//付款方式
				$worksheet->setCellValueByColumnAndRow(11, $row, $result['付款方式']);
				//保证金
				$worksheet->setCellValueByColumnAndRow(12, $row, $result['保证金']);
				//开始接电时间
				$worksheet->setCellValueByColumnAndRow(13, $row, $result['开始接电时间']);
				//电费结算期
				$worksheet->setCellValueByColumnAndRow(14, $row, $result['电费结算期']);
				//电费
				$worksheet->setCellValueByColumnAndRow(15, $row, $result['电费']);
				//电费发票
				$worksheet->setCellValueByColumnAndRow(16, $row, $result['电费发票']);
				//电费支付
				$worksheet->setCellValueByColumnAndRow(17, $row, $result['电费付款']);
				//终止备注
				$worksheet->setCellValueByColumnAndRow(18, $row, $result['终止备注']);
			}
		}
		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
		$writer->save('/Users/kingsley/PhpstormProjects/scanpdf/public/铁通合同100份-合同信息截取测试.xlsx');
	}
	//提取铁通合同数据
	public function getTieTong($content)
	{
		$content=trim($content);
		//地址
		preg_match("/租赁标的采用以下第(.*?)种方式/", $content, $match);
		$select = $match[1];
		if ($select == '(1.1)') {
			preg_match("/{$select}甲方向乙方提供位于(.*?)以及各楼层天花和管井/", $content, $match);
			if(empty($match)){
				preg_match("/{$select}甲方向乙方提供位于(.*?),?供乙方设置移动电话分布系统/", $content, $match);
			}
		} else if ($select == '(1.2)') {
			preg_match("/{$select}甲方向乙方提供位于(.*?)用作乙方建设无人值守/", $content, $match);
		} else if ($select == '(1.3)') {
			preg_match("/③[^)].*?([0-9]+(\.[0-9]{1,4})?)元\/度/", $content, $match);
		}else if ($select == '(1.4)') {
			preg_match("/{$select}甲方向乙方提供位于(.*?)的场地用作乙方自行安装室外天线设备/", $content, $match);
		}
		$result['地址'] = $match[2];
		//租赁期限
		preg_match("/行期间自(.*?)止/", $content, $match);
		$result['租赁期限'] = $match[1];
		//场地开始时间
		preg_match("/因乙方从(.*?)开始使用甲方场地/", $content, $match);
		if(empty($match)){
			$result['场地开始使用时间'] = '';
		}else{
			$result['场地开始使用时间'] = $match[1];
		}
		//交租金日
		preg_match("/首期满后,乙方应于(.*?)向甲方支付第二年/", $content, $match);
		$result['交租金日'] = $match[1];
		//租金
		preg_match("/年租金小计.*?%(.*?)\.2租赁期/", $content, $match);
		if(empty($match)){
			preg_match("/年场地租金(.*?)元\/年/", $content, $match);
			$result['租金'] = str_replace(",",".",$match[1]);
			if(!strstr($result['租金'],'.')){
				$result['租金'] = (intval($result['租金']) / 100);
			}
		}else {
			preg_match("/\.(.*)/", $match[1], $match1);
			$result['租金'] = substr($match1[1], 2);
			preg_match('/-?\\d+(?:\\.\\d+)?/m', $result['租金'], $arr);
			$result['租金'] = sprintf("%.2f", $arr[0]);
		}
		//付款方式
		preg_match("/甲方向乙方开具(.*?)。/", $content, $match);
//		$result['付款方式'] = $match[1];
		$result['付款方式'] = '符合国家法律法规和标准的增值税专用发票';
		//合同保证金
		preg_match("/向甲方交纳人民币(.*?)作为保证金/", $content, $match);
		$result['保证金'] = $match[1];
		//开始接电时间
		preg_match("/从【(.*?)开始,乙方需向甲方按本协议电费标准支付电费/", $content, $match);
		if(empty($match)) {
			$result['开始接电时间'] = '';
		}else{
			$result['开始接电时间'] = str_replace("【", "", $match[1]);
			$result['开始接电时间'] = str_replace("】", "", $result['开始接电时间']);
			if(strpos($result['开始接电时间'],'第4页') ){
				preg_match("/(.*?)第4页/", $result['开始接电时间'], $match);
				$result['开始接电时间'] = $match[1].'日';
			}
		}
		//电费结算期
		preg_match("/电费由乙方与甲方进行结算,(.*?)支付一次/", $content, $match);
		if(empty($match)) {
			$result['电费结算期'] = '';
		}else{
			$result['电费结算期'] = $match[1];
		}
		//是否交电费
		if (preg_match("/费按照以下第1种方/", $content, $match)) {
			$result['电费'] = '直供电';
			$result['电费发票'] = '';
			$result['电费付款'] = '自行与电力局结算';
		}else if (preg_match("/费按照以下第2种方/", $content, $match)) {
			//电价
			preg_match("/电价按(.*?)\(含税\)支付/", $content, $match);
			$result['电费'] = $match[1];
			//电费发票
			preg_match("/甲方先提供(.*?发票)/", $content, $match);
//			$result['电费发票'] = $match[1];
			$result['电费发票'] = '电费增值税专用发票';
			//电费付款
			preg_match("/乙方收到甲方提供的发票后,(.*?电费)/", $content, $match);
//			$result['电费付款'] = $match[1];
			$result['电费付款'] = '在30个工作日内通过银行转账方式支付电费';
		}else{
			$result['电费'] = '免费';
			$result['电费发票'] = '';
			$result['电费付款'] = '';
		}
		//终止备注
		preg_match("/本协议终止,(.*?原状)/", $content, $match);
		if (isset($match[1])) {
			$result['终止备注'] = $match[1];
		}else{
			$result['终止备注'] ='';
		}
		//合同编号
		preg_match("/[s,5]z-?(.*?)合同/i", $content, $match);
		if(empty($match)){
			dump($result);
			throw new Exception('合同编号无法识别，跳过');
		}else{
			$result['合同编号']=$match[1];
			$result['合同编号'] = str_replace("-","",$match[1]);
			$result['合同编号'] = preg_replace('/([\x80-\xff]*)/i','',$result['合同编号']);
			$result['合同编号'] = preg_replace('/([a-zA-Z]*)/','',$result['合同编号']);
			//不够9位在字符串年份后面插入0
		    if(strlen($result['合同编号'])<9){
		    	$len = 9-strlen($result['合同编号']);
			    for ($x=0; $x<$len; $x++) {
				    $result['合同编号']=  substr_replace($result['合同编号'],'0',4,0);
			    }
		    }
		}
		dump($result);
		return $result;
	}
	//提取自有合同
	public function getField()
	{
		$content = file_get_contents('/Users/kingsley/PhpstormProjects/scanpdf/runtime/storage/爱华大厦（L）租赁协议（2016.9.1-2017.8.31）.pdf/爱华大厦（L）租赁协议（2016.9.1-2017.8.31）.pdf.txt');
		//租金价格
		preg_match("/乙方支付甲方的场地.*每年.*?(\d+)元/", $content, $match);
		$result['rent'] = $match[1];
		//租金票据
		preg_match("/向乙方开具(.*?)[，。]/u", $content, $match);
		$result['rent_bill'] = $match[1];
		//电费价格
		preg_match("/用电类型为\((.*?)\)/", $content, $match);
		$select = $match[1];
		if ($select == '①') {
			preg_match("/①[^)].*?([0-9]+(\.[0-9]{1,4})?)元\/度/", $content, $match);
			dump($match);
		} else if ($select == '②') {
			preg_match("/②[^)].*?([0-9]+(\.[0-9]{1,4})?)元\/度/", $content, $match);
			dump($match);
		} else if ($select == '③') {
			preg_match("/③[^)].*?([0-9]+(\.[0-9]{1,4})?)元\/度/", $content, $match);
		}
		$result['power_rate'] = $match[1];
		//电费票据
		preg_match("/向乙.?出具(.*?)[，。]/u", $content, $match);
		$result['power_bill'] = $match[1];
		dump($result);
	}
	//调用ocr提取文字
	public function getString($path)
	{
		try {
			$client = new AipOcr(env('ocr.app_id'), env('ocr.api_key'), env('ocr.secret_key'));
			$string = '';
			foreach ($path as $img) {
				//获取图片信息
				$image                       = file_get_contents($img);
				if($img==$path[1]||$img==$path[2]||$img==$path[10]||$img==$path[11]){
					continue;
				}else if($img==$path[0]||$img==$path[9]||$img==$path[8]||$img==$path[7]){
					$result                      = $client->basicGeneral($image);
				}else {
					$result                      = $client->basicAccurate($image);
				}
				if (array_key_exists('error_code', $result)) {
					throw new Exception('识别不成功' . $result['error_msg']);
				}
				$result = $result['words_result'];
				foreach ($result as $k => $v) {
					$string .= $v['words'];
				}
			}
		} catch (Exception $e) {
			throw $e;
		}
		return $string;
	}
	//表格ocr测试
	public function tableTest(Request $request){
		$file = $request->file('image');
		$client = new AipOcr(env('ocr.app_id'), env('ocr.api_key'), env('ocr.secret_key'));
		$image                       = file_get_contents('/Users/kingsley/PhpstormProjects/scanpdf/runtime/storage/渔景机房等108个站点租赁协议（2020.3.1-20206.30）.pdf/渔景机房等108个站点租赁协议（2020.3.1-20206.30）.pdf(10).png');
		$options = array();
		$options['is_sync']='true';
		$options['request_type']='json';
		$result                      = $client->tableRecognitionAsync($image,$options);
		dump($result);
	}
	//pdfparser
	public function pdfchange()
	{
		$parser = new \Smalot\PdfParser\Parser();
		$pdf    = $parser->parseFile('document.pdf');
		
		$text = $pdf->getText();
		echo $text;
	}
}
