# php-officeplus
EXCEL加强版本，基于phpexcel控件二次开发，支持单元格智能合并，单元格样式自定义。
### 使用限制：
* 类基于phpoffice/excel第三方扩展实现（除CSV外）
* 使用本类需要满足保证构造函数中实例化PHPExcel方法，需要保证构造函数中以下方法调用正常

```
public function __construct()
{
	//从新设置缓存路径
	$cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;
	$cacheSettings = [];
	\PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
	//实例化类
	$this->objPHPExcel = new \PHPExcel();
}

```
* 使用csv导出功能不需要满足上述条件

### 使用方法
* CTable实例化使用导出复杂的单元格数据

```
/* excel通用导出类
 | —————————————————
 | 使用场景
 | ——————————————————
 | 特殊的合并单元
 | 多个sheet导出
 | 一个单元格输出多行数据
 | 单元格样式有要求（行高、颜色、对齐方式等）
 | ——————————————————
 | 如果需要导出csv格式数据
 | 或者简单的导出excel数据
 | 请直接使用static方法
 | ——————————————————
 */
```

* 使用

```
$arRet = [
	'columns' => [],   //表头 
	'datas'   => [],   //数据
];

$table = new CTable();
//CTable::MERGE_ROW_TYPE导出类型，通用类型（推荐使用此类型）
//test_sheet工作区名称，可以多次addSheet创建多个工作区导出
$table->addSheet($arRet, CTable::MERGE_ROW_TYPE, 'test_sheet');
//导出数据
$table->outExcel('test_name');

```

* datas数据格式（二维数组）

```
{
	"datas":[
		{
			"id":1,
			"region":"华东",
			"province":"福建",
			"name":"电信-广西-南宁-节点-1(4/10)",
			"min":10,
			"max":100,
			"region2":"测试"
			"remark":"无"
		},
		{
			"id":2,
			"region":"华东",
			"province":"福建",
			"name":"电信-广西-南宁-节点-2(4/10)",
			"min":10,
			"max":100,
			"region2":"测试"
			"remark":"无"
		},
		{
			"id":3,
			"region":"华东",
			"province":"福建",
			"name":"电信-广西-南宁-节点-3(4/10)",
			"min":10,
			"max":100,
			"region2":"测试"
			"remark":"无"
		},
		{
			"id":4,
			"region":"华东",
			"province":"福建",
			"name":"电信-广西-南宁-节点-4(4/10)",
			"min":10,
			"max":100,
			"region2":"测试"
			"remark":"无"
		}
	]
}
```

* columns格式
	* 基础格式，所有子项为展示项，未配置columns的列不会写入excel，title：表头名，key：对应每个子项的数据key值

	```
	{
		"columns":[
			{
				"title":"节点id",
				"key":"id"
			},
			{
				"title":"省份",
				"key":"province"
			},
			{
				"title":"节点",
				"key":"name"
			}
		]
	}
	```
	
	* 表头样式

	```
	{
	    "columns":[
	        {
	            "title":"节点id",
	            "key":"id",
	            "style":{
	            	"width":"auto",     //宽度自适应（"auto"或者int）。整列
	            	"color":"000000",   //字体颜色（RGB编码，无#）。本单元格
	            	"align":"right",    //对齐方式，left、center、right。本单元格
	            	"height":50         //行高。标题行
	            }
	        },
	        {                          //表头合并，keys所有子项的表头合并起来，共用"所在地区"
	            "title":"所在地区",
	            "keys":[
	            	{
	            		"key":"region",
	            		"style":{
	            			"width":50
	            		}
	            	},
	            	{
	            		"key":"province",
	            		"style":{
	            			"width":"auto"
	            		}
	            	}
	            ]
	        },
	        {
	            "title":"节点",
	            "key":"name"
	        }
      	]
	}
	```

	* 行合并

	```
	{
		"columns":[
			{
				"title":"节点id",
				"key":"id"
			},
			{
				"title":"大区",
				"key":"region",
				"merge":true,      //列书否合并，合并会将同列所有相同数据进行聚合合并，变成横向的多级子树样式
				"sort":0,  			 //合并优先级，写的优先于不写的，不写的按照位置先后顺序合并
				"affect":[         //相同的合并策略列，数组中的列会根据本列合并情况做相同的合并
					"id",
					"region2"
				]
			},
			{
				"title":"省份",
				"key":"province",
				"merge":true,
				"sort":1
			},
			{
				"title":"节点",	
				"key":"name"
			},
			{
				"title":"大区2",
				"key":"region2",
			}
		]
	}
	
	```
	
	* 自定义单元格样式vender回调函数

	```
	//特殊的单元格处理，col为当前单元格列，row当前行，data本行的数据
	$vendor = function ($col, $row, $data) {
		if ($data['name'] == '电信-广东-佛山-节点-1(5/0)') {
			return [
				'merge'     => [  //合并单元格，没有合并则不返回此项
					$row,				 //开始行
					$col, 			 //开始列
					$row,				 //结束行
					$col + 1,       //结束列
				],
				'primitive' => [   //使用原生格式设置方法，设置此项则style失效
				],
				'style'     => [   //修改本单元格样式，参照表头的style格式
					'color'  => 'A52A2A',
					'align'  => 'right',
					'height' => 50
				],
			];
		}
		
		return [];     //无返回则不进行任何操作
	};
	
	{
		"columns":[
			{
				"title":"节点id",
				"key":"id"
			},
			{
				"title":"省份",
				"key":"province"
			},
			{
				"title":"节点",
				"key":"name",
				"vendor":$vendor     //此列引入特殊配置的回调函数
			}
		]
	}

	```
	
### 数据简单导出
* PHPExcel本身具有性能限制，大数量的数据可能会引起内存跑满
* 在数据量巨大（十万百万级以上）的数据导出时（例如导出一个月所有节点的每个时刻带宽数据，大概100万条），建议使用csv格式导出
* 使用方法

```
//columns只支持基础类型
CTable::outCsv('test_name', $columns, $datas);
```




