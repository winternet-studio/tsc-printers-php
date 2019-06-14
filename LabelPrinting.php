<?php
namespace winternet\tscprinters;

/**
 * Library for handling TSC label printers
 *
 * If images are used PHP's GD extension is required.
 * If USB connection is used PHP's COM (com_dotnet) extension is required.
 * If HTTP-UNOFFICIAL connection method is used PHP's curl extension is required, unless you implement your own HTTP client.
 *
 * A TSC DA220 printer has been used in the development of this class.
 *
 * Example:
 * ```
 * $tsc = new \winternet\tscprinters\LabelPrinting(['USB', 'TSPL', 'TSCDA220'], ['debug' => true]);
 * $tsc->newLabel([
 * 	'width' => 101.6,
 * 	'height' => 152.4,
 * 	'sensor' => 'gap',
 * 	'vertical' => 3,
 * ]);
 * $tsc->addText(['text' => '123000003', 'font' => '2', 'x' => 0, 'y' => 0]);
 * $tsc->addText(['text' => 'ABCDefghi', 'font' => '2', 'x' => 0, 'y' => 30]);
 * $tsc->addText(['text' => '987650001', 'font' => ['type' => 'windowsfont', 'facename' => 'Arial', 'fontheight' => 48, 'fontstyle' => 0, 'fontunderline' => 0], 'x' => 0, 'y' => 60]);
 * $tsc->addLine(['x' => 0, 'y' => 60, 'width' => 30, 'thickness' => 3]);
 * $tsc->addBarcode(['x' => 0, 'y' => 70, 'barcodeType' => '128', 'height' => 50, 'data' => 'SW1234578']);
 * $tsc->addImage(['imageFile' => __DIR__ .'/myimage.png', 'x' => 0, 'y' => 60]);
 * $tsc->printLabel();
 * var_dump($tsc->debugInfo);
 *
 * // Examples of custom commands:
 * $tsc->customCommand('GAPDETECT');  //calls ActiveXsendcommand()
 * $tsc->customCommand('DIRECTION 1,0');
 * $tsc->customCommand('PUTPCX 1,1,"UL.PCX"');
 * $tsc->customCommand('PUTBMP 10,10,"UL.BMP"');
 * $tsc->customCommand('FILES');
 * $tsc->customCommand('KILL F,"UL.BMP"');
 * $tsc->callActiveX('downloadpcx', 'C:\myfiles\UL.PCX', 'UL.PCX');  //calls ActiveXdownloadpcx()
 * $tsc->callActiveX('windowsfont', 400, 200, 48, 0, 3, 1, 'arial', 'DEG 0');  //calls ActiveXwindowsfont()
 * $tsc->callActiveX('windowsfont', 400, 200, 48, 90, 3, 1, 'arial', 'DEG 90');
 * $tsc->callActiveX('windowsfont', 400, 200, 48, 180, 3, 1, 'arial', 'DEG 180');
 * $tsc->callActiveX('windowsfont', 400, 200, 48, 270, 3, 1, 'arial', 'DEG 270');
 * $tsc->callActiveX('barcode', '100', '40', '128', '50', '1', '0', '2', '2', '123456789');
 * ```
 *
 * ZPL basic example:
 * ```
 * ^XA
 * ^CI28
 * ^FO100,100^A0N,50,50^FDEspañol^FS
 * ^FO0,0^A0N,50,50^FDWinterNet Studio^FS
 * ^FO50,50^A0N,50,50^FDNorway^FS
 * ^FO150,150^A0N,50,50^FDIn God We Trust^FS
 * ^XZ
 * ```
 */
class LabelPrinting {

	/**
	 * @var array : For holding connection configuration
	 */
	public $connectConfig;

	/**
	 * @var string : Printer IP address when connected over the network
	 */
	public $ipAddress;

	/**
	 * @var array : For holding misc options - see constructor
	 */
	public $options = [];

	/**
	 * @var object : For holding the instantiated COM object
	 */
	public $com = null;

	/**
	 * @var string : Whether we are using `ZPL` or `TSPL` as printer language
	 */
	public $language = null;

	/**
	 * @var string : Whether we are using `USB` or `HTTP-UNOFFICIAL` interface
	 */
	public $interface = null;

	/**
	 * @var array : Buffer for holding commands for writing label
	 */
	public $commandBuffer = [];

	/**
	 * @var boolean : Whether a new label has been initialized
	 */
	public $labelInitialized = false;

	/**
	 * @var array : Filled with debug info if debug is enabled
	 */
	public $debugInfo = [];

	protected $brightnessThreshold = null;

	/**
	 * @param array $connectConfig : Configuration for connecting to the printer. Examples:
	 *   - `['USB', 'ZPL', 'TSCDA220']` for connecting via USB
	 *       - requires PHP's COM extension (com_dotnet) as well as these installations:
	 *           - register TscActiveX.dll according to document in PHP SDK example for the printer (eg. https://www.tscprinters.com/EN/support/support_download/DA210-DA220%20Series)
	 *           - download and install printer driver from TSC website (not the Diagnostic Tool)
	 *           - plugin the printer when it asks you to and wait for it to detect it
	 *       - the second parameter is the printer language to use, either `TSPL` or `ZPL` (TSPL is relatively slow for images since we are required to use a work-around)
	 *       - the third parameter is the printer name you set during driver installation (NOT the share name, nor the USB port number (eg. "USB001"))
	 *   - `['HTTP-UNOFFICIAL', 'ZPL', '192.168.1.100']´ for connecting purely over ethernet (sort of like a REST API)
	 *       - using the printer's unofficial web interface to upload a file with ZPL commands. Doesn't require the printer driver, COM .dll registration, or anything else installed.
	 *       - the second parameter is the printer language to use: `ZPL` or `TSPL`
	 *   - Notes: There is no guarantee that using ZPL vs TSPL will render exactly the same way, so don't expect that you can just switch language without adjusting parameters in your method calls.
	 * @param array $options : Available options:
	 *   - `httpHandler` : option for implementing your own HTTP client to handle the process of communicating with the unofficial web interface of the printer. See code for arguments and details.
	 *   - `debug` : set true to enable debugging (eg. include ZPL commands in returned array from `printLabel()`)
	 */
	function __construct($connectConfig, $options = []) {
		$this->connectConfig = $connectConfig;
		$this->options = $options;

		if ($connectConfig[1] === 'TSPL') {
			$this->language = 'TSPL';
		} elseif ($connectConfig[1] === 'ZPL') {
			$this->language = 'ZPL';
		} else {
			throw new \Exception('Invalid second parameter of connectConfig in the TscPrinter constructor.');
		}

		if ($connectConfig[0] === 'USB') {
			if (!extension_loaded('com_dotnet')) {
				throw new \Exception('PHP COM extension is not installed.');
			}

			$this->interface = 'USB';
			if ($connectConfig[2]) {
				$this->com = new \COM('TSCActiveX.TSCLIB');
				if ($this->options['debug']) {
					$this->debugInfo[] = ['comObject' => gettype($this->com)];
				}
				$this->callActiveX('openport', $connectConfig[2]);
			} else {
				throw new \Exception('Missing third parameter of connectConfig in the TscPrinter constructor.');
			}

		} elseif ($connectConfig[0] === 'HTTP-UNOFFICIAL') {
			$this->interface = 'HTTP-UNOFFICIAL';
			$this->ipAddress = $connectConfig[2];

		} else {
			throw new \Exception('Invalid first parameter of connectConfig in the TscPrinter constructor.');
		}
	}

	function __destruct() {
		if ($this->interface === 'USB') {
			$this->callActiveX('closeport');
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - ...
	 *   - `sensor` : `gap` or `blackMark`
	 *   - `characterSet` : 
	 *       ZPL: character set according to page 151 in "ZPL II Programming Guide Vol 1.pdf" (Part Number: 13979L-010 Rev. A, dated 6/19/08)
	 *       TSPL: codepage according to page 30 in "TSPL_TSPL2_Programming.pdf" from https://www.tscprinters.com/EN/support/support_download/DA210-DA220%20Series
	 *   - ...
	 */
	public function newLabel($params = []) {
		$params = array_merge([
			'width' => 101.6,
			'height' => 152.4,
			'speed' => 5,
			'density' => 8,
			'sensor' => 'gap',
			'vertical' => 3,
			'offset' => 0,
			'characterSet' => null,
		], $params);

		$this->commandBuffer = [];

		if ($this->language === 'TSPL') {
			if ($this->interface === 'USB') {
				$this->callActiveX('setup', $params['width'], $params['height'], $params['speed'], $params['density'], ($params['sensor'] === 'gap' ? 0 : 1), $params['vertical'], $params['offset']);
				$this->callActiveX('clearbuffer');
				if ($params['characterSet']) {
					$this->callActiveX('sendcommand', 'CODEPAGE '. $params['characterSet']);
				}
			} elseif ($this->interface === 'HTTP-UNOFFICIAL') {
				// TODO: set label size etc

			}

		} elseif ($this->language === 'ZPL') {
			// TODO: set label size etc

			$this->commandBuffer[] = '^XA';
			if ($params['characterSet']) {
				$this->commandBuffer[] = '^CI'. $params['characterSet'];
			} else {
				$this->commandBuffer[] = '^CI28';  //enable unicode
			}

			if ($this->interface === 'USB') {
				$this->executeCommandBuffer();
			}
		}

		$this->labelInitialized = true;
	}

	/**
	 * @param array $params : Available parameters:
	 *   - `x`
	 *   - `y`
	 *   - `text` : the text to write
	 *   - `font` : the font to use
	 *       - ZPL: see page 894 in "ZPL II Programming Guide Vol 1.pdf" (Part Number: 13979L-010 Rev. A, dated 6/19/08)
	 *       - TSPL: see page 90 in "TSPL_TSPL2_Programming.pdf" from https://www.tscprinters.com/EN/support/support_download/DA210-DA220%20Series
	 *         - to use the ActiveX functions ActiveXprinterfont() and ActiveXwindowsfont() mentioned in "ActiveX Dll Functions Description.doc" you can provide arrays like this:
	 *           - `['type' => 'printerfont', 'font' => '3']`
	 *           - `['type' => 'windowsfont', 'facename' => 'Arial', 'fontheight' => 48, 'fontstyle' => 0, 'fontunderline' => 0]`
	 *   - `xMultiplication` : 1-10 (TSPL only)
	 *   - `yMultiplication` : 1-10 (TSPL only)
	 *   - `characterHeight` : scalable font: 10-32000, bitmapped font: 1-10 (ZPL only)
	 *   - `characterWidth`  : scalable font: 10-32000, bitmapped font: 1-10 (ZPL only)
	 *   - `alignment` : `left` or `center` or `right` (TSPL only)
	 *   - `rotation` : 0 or 90 or 180 or 270
	 */
	public function addText($params) {
		$params = array_merge([
			'x' => null,
			'y' => null,
			'text' => '',
			'font' => '0',
			'xMultiplication' => ($this->language === 'TSPL' && (string) $params['font'] === '0' ? 8 : 1),  //For font "0", this parameter is used to specify the width (point) of true type font. 1 point=1/72 inch
			'yMultiplication' => ($this->language === 'TSPL' && (string) $params['font'] === '0' ? 8 : 1),  //For font "0", this parameter is used to specify the height (point) of true type font. 1 point=1/72 inch
			'characterHeight' => 30,
			'characterWidth' => 30,
			'alignment' => 'left',
			'rotation' => 0,
		], $params);

		$this->checkInit();
		if ($this->language === 'TSPL') {
			$alignMap = [
				'left' => '1',
				'center' => '2',
				'right' => '3',
			];
			if ($this->interface === 'USB' && is_array($params['font']) && $params['font']['type'] === 'printerfont') {
				$this->callActiveX('printerfont', $params['x'], $params['y'], $params['font']['font'], $params['rotation'], $params['xMultiplication'], $params['yMultiplication'], $params['text']);
			} elseif ($this->interface === 'USB' && is_array($params['font']) && $params['font']['type'] === 'windowsfont') {
				$this->callActiveX('windowsfont', $params['x'], $params['y'], $params['font']['fontheight'], $params['rotation'], $params['font']['fontstyle'], $params['font']['fontunderline'], $params['font']['facename'], $params['text']);
			} else {
				if (is_array($params['font'])) {
					throw new \Exception('The font parameter may only be an array when USB interface is used (only when COM object and TSPL is used).');
				}
				$tspl = 'TEXT '. $params['x'] .','. $params['y'] .',"'. str_replace('"', "\\\"", (string) $params['font']) .'",'. $params['rotation'] .','. $params['xMultiplication'] .','. $params['yMultiplication'] .','. $alignMap[$params['alignment']] .',"'. str_replace('"', "\\\"", (string) $params['text']) .'"';

				if ($this->interface === 'USB') {
					$this->callActiveX('sendcommand', $tspl);

				} elseif ($this->interface === 'TSPL') {
					$this->commandBuffer[] = $tspl;
				}
			}

		} elseif ($this->language === 'ZPL') {
			if (is_array($params['font'])) {
				throw new \Exception('The font parameter may not be an array when ZPL is used (only when COM object and TSPL is used).');
			}

			$rotateMap = [
				0 => 'N',
				90 => 'R',
				180 => 'I',
				270 => 'B',
			];
			$zpl  = '^FO'. $params['x'] .','. $params['y'] .'^A'. $params['font'] .','. $rotateMap[$params['rotation']] .','. $params['characterHeight'] .','. $params['characterWidth'];
			$zpl .= '^FD'. $params['text'] .'^FS';

			$this->commandBuffer[] = $zpl;

			if ($this->interface === 'USB') {
				$this->executeCommandBuffer();
			}
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - ...
	 *   - `barcodeType` : see TSPL and ZPL documentation. Eg. `128` in TSPL.
	 *   - `height` : barcode height in dots
	 *   - `printPlaintext` : `below`, `above` or `none` (`above` only supported in ZPL)
	 *   - `plaintextAlignment` : `left` or `center` or `right` (TSPL only)
	 *   - `rotation` : 0 or 90 or 180 or 270
	 *   - `narrowWidth` : width of narrow element (in dots)
	 *   - `wideWidth` : width of wide element (in dots)
	 *   - `alignment` : `left` or `center` or `right` (TSPL only)
	 *   - `data` : data to make barcode of
	 */
	public function addBarcode($params) {
		$params = array_merge([
			'x' => null,
			'y' => null,
			'barcodeType' => '128',
			'height' => 50,
			'printPlaintext' => 'below',
			'plaintextAlignment' => 'center',
			'rotation' => 0,
			'narrowWidth' => 3,
			'wideWidth' => 5,
			'alignment' => 'left',
			'data' => '',
		], $params);

		$this->checkInit();
		if ($this->language === 'TSPL') {
			if ($params['printPlaintext'] === 'none') {
				$humanReadable = 0;
			} else {
				if ($params['plaintextAlignment'] === 'left') {
					$humanReadable = 1;
				} elseif ($params['plaintextAlignment'] === 'center') {
					$humanReadable = 2;
				} elseif ($params['plaintextAlignment'] === 'right') {
					$humanReadable = 3;
				}
			}
			$alignMap = [
				'left' => '1',
				'center' => '2',
				'right' => '3',
			];
			$tspl = 'BARCODE '. $params['x'] .','. $params['y'] .',"'. str_replace('"', "\\\"", (string) $params['barcodeType']) .'",'. $params['height'] .','. $humanReadable .','. $params['rotation'] .','. $params['narrowWidth'] .','. $params['wideWidth'] .','. $alignMap[$params['alignment']] .',"'. str_replace('"', "\\\"", (string) $params['data']) .'"';

			$this->commandBuffer[] = $tspl;

			/*
			ActiveX examples - but they have less features:
			$this->callActiveX('barcode', '50', '100', '128', '70', '0', '0', '3', '1', '123456');
			$this->callActiveX('barcode', '100', '40', '128', '50', '1', '0', '2', '2', '123456789');
			*/

		} elseif ($this->language === 'ZPL') {
			// TODO

			$this->commandBuffer[] = '...';

		}

		if ($this->interface === 'USB') {
			$this->executeCommandBuffer();
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - `cornerRadius` : ZPL: 0-8, TSPL: 0-?
	 *   - `borderColor` : ZPL: `black` or `white`, TSPL: <not available>
	 */
	public function addLine($params, $executeImmediately = true) {
		// Apply thickness when using TSPL by setting height of box
		if ($this->language === 'TSPL') {
			if (array_key_exists('thickness', $params) && !array_key_exists('width', $params)) {
				$params['width'] = $params['thickness'];
			} elseif (array_key_exists('thickness', $params) && !array_key_exists('height', $params)) {
				$params['height'] = $params['thickness'];
			}
		}
		$params = array_merge([
			'x' => null,
			'y' => null,
			'width' => ($this->language == 'ZPL' ? 0 : 1),
			'height' => ($this->language == 'ZPL' ? 0 : 1),
			'thickness' => 1,
			'cornerRadius' => 0,
			'borderColor' => 'black',
		], $params);

		$this->checkInit();
		if ($this->language === 'TSPL') {
			$this->commandBuffer[] = 'BAR '. $params['x'] .','. $params['y'] .','. $params['width'] .','. $params['height'];

		} elseif ($this->language === 'ZPL') {
			$this->commandBuffer[] = '^FO'. $params['x'] .','. $params['y'] .'^GB'. $params['width'] .','. $params['height'] .','. $params['thickness'] .','. strtoupper(substr($params['borderColor'], 0, 1)) .','. $params['cornerRadius'] .'^FS';
		}

		if ($this->interface === 'USB' && $executeImmediately) {
			$this->executeCommandBuffer();
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - `cornerRadius` : ZPL: 0-8, TSPL: 0-?
	 *   - `borderColor` : ZPL: `black` or `white`, TSPL: <not available>
	 */
	public function addBox($params) {
		$params = array_merge([
			'ulX' => null,
			'ulY' => null,
			'lrX' => null,
			'lrY' => null,
			'thickness' => 1,
			'cornerRadius' => 0,
			'borderColor' => 'black',
		], $params);

		$this->checkInit();
		if ($this->language === 'TSPL') {
			$this->commandBuffer[] = 'BOX '. $params['ulX'] .','. $params['ulY'] .','. $params['lrX'] .','. $params['lrY'] .','. $params['thickness'] .','. $params['cornerRadius'];

		} elseif ($this->language === 'ZPL') {
			$this->commandBuffer[] = '^FO'. $params['ulX'] .','. $params['ulY'] .'^GB'. ($params['lrX'] - $params['ulX']) .','. ($params['lrY'] - $params['ulY']) .','. $params['thickness'] .','. strtoupper(substr($params['borderColor'], 0, 1)) .','. $params['cornerRadius'] .'^FS';
		}

		if ($this->interface === 'USB') {
			$this->executeCommandBuffer();
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - `x` : X position
	 *   - `y` : Y position
	 *   - `imageFile` : full path to image file (currently only PNG is supported)
	 *   - `imageFileIsContent` : set true if you provide the file content in `imageFile` instead of a path to a file
	 */
	public function addImage($params) {
		$params = array_merge([
			'x' => null,
			'y' => null,
			'imageFile' => null,
			'imageFileIsContent' => false,
		], $params);

		if (!file_exists($params['imageFile'])) {
			throw new \Exception('Image file '. $params['imageFile'] .' does not exist.');
		}

		$this->checkInit();
		if ($this->language === 'TSPL') {
			if (!extension_loaded('gd')) {
				throw new \Exception('PHP gd extension is not installed.');
			}

			if (!$params['imageFileIsContent'] && !file_exists($params['imageFile'])) {
				throw new \Exception('Image file to print line by line does not exist.');
			}

			// This is a work-around for printing images which is normally not supported in TSPL within PHP
			if ($params['imageFileIsContent']) {
				$im   = imagecreatefromstring($params['imageFile']);
				$size = getimagesizefromstring($params['imageFile']);
			} else {
				$im   = imagecreatefrompng($params['imageFile']);
				$size = getimagesize($params['imageFile']);
			}
			$width  = $size[0];
			$height = $size[1];

			$xPos = 0;  //dots, negative values are allow (1 mm = 8 dots)
			$yPos = 0;  //dots, negative values are allow (1 mm = 8 dots)
			$brightnessThreshold = $this->getBrightnessThreshold();

			$printLine = function($startXindex, $endXindex, $yIndex) use (&$xPos, &$yPos) {
				$width = $endXindex - $startXindex;
				$this->addLine(['x' => $startXindex+1 + $xPos, 'y' => $yIndex+1 + $yPos, 'width' => $width, 'height' => 1], false);
			};

			// $o = '';
			for ($y = 0; $y < $height; $y++) {
				$startX = null;
				$endX = null;
				for ($x = 0; $x < $width; $x++) {
					$rgb = imagecolorat($im, $x, $y);

					// Change from zero-based index to normal number
					// $x++;
					// $y++;

					$r = ($rgb >> 16) & 0xFF;
					$g = ($rgb >> 8) & 0xFF;
					$b = $rgb & 0xFF;
					// var_dump($r, $g, $b);
					// $alpha = ($rgb & 0x7F000000) >> 24;
					// var_dump($alpha);

					$brightness = $this->calculateBrightness($r, $g, $b);
					$isBlack = ($brightness <= $brightnessThreshold ? true : false);

					if ($isBlack && $startX === null) {
						// starting a new line with black
						$startX = $x;
					} elseif (!$isBlack && $startX !== null) {
						// we have come to the end of a line with black
						$printLine($startX, $x, $y);
						$startX = null;
					}
					// $o .= $x .','. $y .' : '. $brightness . PHP_EOL;
					// echo $brightness;
					// echo '<br>';
				}

				if ($startX !== null) {  //print line that goes all the way to the right edge
					$printLine($startX, $x, $y);
					$startX = null;
				}
			}
			// file_put_contents('o.txt', $o);

		} elseif ($this->language === 'ZPL') {
			// TODO
			$grf = $this->imageToGrf($params['imageFile']);
			$zpl = '^FO'. $params['x'] .','. $params['y'] .'^GFA,'. $grf['bytesTotal'] .','. $grf['bytesTotal'] .','. $grf['bytesPerRow'] .','. $grf['hexString'] .'^FS';
			// TODO: this doesn't work yet!

			$this->commandBuffer[] = $zpl;
		}

		if ($this->interface === 'USB') {
			$this->executeCommandBuffer();
		}
	}

	/**
	 * @param array $params : Available parameters:
	 *   - `labelSets` : number of label sets, default 1
	 *   - `copies` : number of copies, default 1
	 * @return mixed : Probably only meaningful for the HTTP-UNOFFICIAL interface
	 */
	public function printLabel($params = []) {
		$params = array_merge([
			'labelSets' => 1,
			'copies' => 1,
		], $params);

		$this->checkInit();
		if ($this->language === 'TSPL') {
			if ($this->interface === 'USB') {
				$result = $this->callActiveX('printlabel', $params['labelSets'], $params['copies']);
			} elseif ($this->interface === 'HTTP-UNOFFICIAL') {
				// TODO
			}

		} elseif ($this->language === 'ZPL') {
			$this->commandBuffer[] = '^XZ';

			if ($this->interface === 'USB') {
				$this->executeCommandBuffer();

			} elseif ($this->interface === 'HTTP-UNOFFICIAL') {
				$commandsString = implode(PHP_EOL, $this->commandBuffer);
				$this->commandBuffer = [];  //clear line buffer

				if (is_callable($this->options['httpHandler'])) {
					$result = $this->options['httpHandler']($this->ipAddress, $commandsString);
				} else {
					$result = $this->httpUploadZplFile($this->ipAddress, $commandsString);
				}

				if ($this->options['debug']) {
					$this->debugInfo[] = $commandsString;
					$result['zplCommands'] = $commandsString;
				}
			}
		}

		return $result;
	}

	/**
	 * @param string $command : Any TSPL or ZPL command, depending on your connection configuration
	 */
	public function customCommand($command) {
		$this->commandBuffer[] = $command;

		if ($this->interface === 'USB') {
			$this->executeCommandBuffer();
		}
	}

	/**
	 * Call an ActiveX function on the COM object
	 *
	 * @param string $function,... : Which function to call. See ActiveX DLL documentation from TSC. Eg. `openport`, `setup`, `sendcommand`, `clearbuffer`, `printlabel` etc.
	 *   - all following arguments will be passed on the ActiveX function.
	 */
	public function callActiveX($function) {
		$passOnArguments = func_get_args();
		unset($passOnArguments[0]);

		// NOT NEEDED
		$passOnArguments = array_map(function($value) {
			return (string) $value;
		}, $passOnArguments);

		$activeXFunction = 'ActiveX'. $function;
		call_user_func_array([$this->com, $activeXFunction], $passOnArguments);

		if ($this->options['debug']) {
			$this->debugInfo[] = ['activeXfunction' => $activeXFunction, 'arguments' => $passOnArguments];
		}
	}

	/**
	 * Execute the commands in the command buffer using the USB/COM object ActiveX calls
	 *
	 * The buffer is emptied afterwards.
	 */
	public function executeCommandBuffer() {
		foreach ($this->commandBuffer as $command) {
			$this->callActiveX('sendcommand', $command);
		}
		$this->commandBuffer = [];
	}

	/**
	 * Calculate the brightness from an RGB value
	 *
	 * @param integer $r : Values in the range 0-255
	 * @param integer $g : Values in the range 0-255
	 * @param integer $b : Values in the range 0-255
	 */
	public function calculateBrightness($r, $g, $b) {
		return 0.2126 * $r + 0.7152 * $g + 0.0722 * $b;  //source: https://stackoverflow.com/questions/596216/formula-to-determine-brightness-of-rgb-color
	}

	/**
	 * Set the brightness threshold
	 */
	public function setBrightnessThreshold($value) {
		return $this->brightnessThreshold = $value;
	}

	/**
	 * Get the brightness threshold
	 *
	 * Defaults to 50% threshold.
	 */
	public function getBrightnessThreshold() {
		if ($this->brightnessThreshold === null) {
			$this->brightnessThreshold = 255 / 2;
		}
		return $this->brightnessThreshold;
	}

	/**
	 * Convert millimeters to dots
	 *
	 * - 203 DPI: 1 mm = 8 dots
	 * - 300 DPI: 1 mm = 11.8 dots
	 * - TSPL: only integer portion will be used. Ex. 2 mm = 23.6 dots then 23 dots will be used, so we round to nearest whole number
	 */
	function mmToDots($mm, $dotsPerMm = 8) {
		return round($mm * $dotsPerMm);
	}

	/**
	 * Converts a PNG image to a .GRF file for use with Zebra printers
	 *
	 * The input is preferably a 1-bit black/white image but RGB images
	 * are accepted as well.
	 *
	 * This function uses PHP's GD library image functions.
	 *
	 * Orignally from https://gist.github.com/thomascube/9651d6fa916124a9c52cb0d4262f2c3f
	 * Received email from Thomas (thomas@brotherli.ch) on 2019-06-12 saying "Yes, the code is free to be used and modified for any purpose but without warranty. It’s just a code snipped I copied there and probably needs refinement. Therefore I didn’t even consider to choose a license. Consider it MIT licensed or similar."
	 * Allan Jensen, WinterNet Studio has made minor modifications to the code.
	 *
	 * @copyright Thomas Bruederli <inbox@brotherli.ch>
	 *
	 * @param string $filename : Path to the input file (or file content is you set `$isFileContent`=true)
	 * @param boolean $isFileContent : Set true if you provide the file content in `$filename` instead of a path to a file.
	 * @return array
	 */
	public function imageToGrf($filename, $isFileContent = false) {
		if (!extension_loaded('gd')) {
			throw new \Exception('PHP gd extension is not installed.');
		}

		if (!$isFileContent && !file_exists($filename)) {
			throw new \Exception('Image file to convert to GRF does not exist.');
		}

		if ($isFileContent) {
			$info = getimagesizefromstring($filename);
			$im = imagecreatefromstring($filename);
		} else {
			$info = getimagesize($filename);
			$im = imagecreatefrompng($filename);
		}
		$width = $info[0]; // imagesx($im);
		$height = $info[1]; // imagesy($im);
		$depth = $info['bits'] ?: 1;
		$threshold = $depth > 1 ? 160 : 0;
		$hexString = '';
		$byteShift = 7;
		$currentByte = 0;
		// iterate over all image pixels
		for ($y = 0; $y < $height; $y++) {
			for ($x = 0; $x < $width; $x++) {
				$color = imagecolorat($im, $x, $y);
				// compute gray value from RGB color
				if ($depth > 1) {
					$value = max($color >> 16, $color >> 8 & 255, $color & 255);
				} else {
					$value = $color;
				}
				// set (inverse) bit for the current pixel
				$currentByte |= (($value > $threshold ? 0 : 1) << $byteShift);
				$byteShift--;
				// 8 pixels filled one byte => append to output as hex
				if ($byteShift < 0) {
					$hexString .= sprintf('%02X', $currentByte);
					$currentByte = 0;
					$byteShift = 7;
				}
			}
			// append last byte at end of row
			if ($byteShift < 7) {
				$hexString .= sprintf('%02X', $currentByte);
				$currentByte = 0;
				$byteShift = 7;
			}
			$hexString .= PHP_EOL;
		}
		$bytesTotal = ceil(($width * $height) / 8);
		$bytesPerRow = ceil($width / 8);

		return [
			'bytesTotal' => $bytesTotal,
			'bytesPerRow' => $bytesPerRow,
			'hexString' => trim($hexString),
		];
	}

	public function checkInit() {
		if (!$this->labelInitialized) {
			throw new \Exception('You need to call newLabel() before you can start adding content to the label.');
		}
	}

	/**
	 * Utility method for uploading file to TSC label printer's unofficial web interface
	 *
	 * @param string $ip : IP address of the printer
	 * @param string $command : ZPL commands
	 */
	public function httpUploadZplFile($ip, $command) {
		$url = 'http://'. $ip .'/admin/cgi-bin/function.cgi';

		$files = [ ['formFieldName' => 'send', 'fileName' => 'zpl.txt', 'fileContent' => $command] ];
		$normalPostFields = [];

		$boundary = uniqid();
		$delimiter = '-------------'. $boundary;
		$postData = $this->buildMime($boundary, $normalPostFields, $files);

		if (!extension_loaded('curl')) {
			throw new \Exception('PHP curl extension is not installed.');
		}

		return $this->httpRequest('POST', $url, $postData, [
			'curlOptions' => [
				CURLOPT_HTTPHEADER => [
					'Content-Type: multipart/form-data; boundary='. $delimiter,
					'Content-Length: '. strlen($postData),
				],
			],
		]);
	}

	/**
	 * Utility method for making HTTP request
	 */
	public function httpRequest($method, $url, $data, $options = []) {
		if (!extension_loaded('curl')) {
			throw new \Exception('PHP curl extension is not installed.');
		}

		$ch = curl_init();
		curl_setopt_array($ch, [
			CURLOPT_URL => $url,
			CURLOPT_RETURNTRANSFER => 1,
			CURLOPT_CUSTOMREQUEST => $method,
			CURLOPT_POST => ($method === 'POST' ? 1 : 0),
			CURLOPT_POSTFIELDS => $data,
		]);
		if (is_array($options['curlOptions'])) {
			foreach ($options['curlOptions'] as $curlopt => $value) {
				curl_setopt($ch, $curlopt, $value);
			}
		}

		$response = curl_exec($ch);

		$info = curl_getinfo($ch);
		$errorMessage = curl_error($ch);

		curl_close($ch);

		return [
			'response' => $response,
			'success' => (strpos($response, 'Send File to Printer') === false ? false : true),
			'errorMessage' => $errorMessage,
			'curlInfo' => $info,
		];
	}

	/**
	 * Utility method for building MIME message
	 */
	public function buildMime($boundary, $normalPostFields, $files) {
		// Source: https://stackoverflow.com/a/45847575/2404541
		$data = '';
		$eol = "\r\n";

		$delimiter = '-------------' . $boundary;

		foreach ($normalPostFields as $name => $content) {
			$data .= '--'. $delimiter . $eol
				. 'Content-Disposition: form-data; name="'. $name .'"'. $eol . $eol
				. $content . $eol;
		}

		foreach ($files as $file) {
			$data .= '--'. $delimiter . $eol
				. 'Content-Disposition: form-data; name="'. $file['formFieldName'] .'"; filename="'. $file['fileName'] .'"'. $eol
				. ($file['fileContentType'] ? 'Content-Type: '. $file['fileContentType'] . $eol : '')
				. 'Content-Transfer-Encoding: binary'. $eol;

			$data .= $eol;
			$data .= $file['fileContent'] . $eol;
		}
		$data .= '--'. $delimiter .'--'. $eol;

		return $data;
	}
}
