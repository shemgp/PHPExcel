<?php

/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */
/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    /**
     * @ignore
     */
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

/**
 * PHPExcel_Reader_HTML
 *
 * @category   PHPExcel
 * @package    PHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Reader_HTML extends PHPExcel_Reader_Abstract implements PHPExcel_Reader_IReader
{

    /**
     * Input encoding
     *
     * @var string
     */
    protected $_inputEncoding = 'ANSI';

    /**
     * Sheet index to read
     *
     * @var int
     */
    protected $_sheetIndex = 0;

    /**
     * Formats
     *
     * @var array
     */
    protected $_formats = array(
        'h1' => array('font' => array('bold' => true,
                'size' => 24,
            ),
        ), //	Bold, 24pt
        'h2' => array('font' => array('bold' => true,
                'size' => 18,
            ),
        ), //	Bold, 18pt
        'h3' => array('font' => array('bold' => true,
                'size' => 13.5,
            ),
        ), //	Bold, 13.5pt
        'h4' => array('font' => array('bold' => true,
                'size' => 12,
            ),
        ), //	Bold, 12pt
        'h5' => array('font' => array('bold' => true,
                'size' => 10,
            ),
        ), //	Bold, 10pt
        'h6' => array('font' => array('bold' => true,
                'size' => 7.5,
            ),
        ), //	Bold, 7.5pt
        'a' => array('font' => array('underline' => true,
                'color' => array('argb' => PHPExcel_Style_Color::COLOR_BLUE,
                ),
            ),
        ), //	Blue underlined
        'hr' => array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array(\PHPExcel_Style_Color::COLOR_BLACK,
                    ),
                ),
            ),
        ), //	Bottom border
    );

    protected $rowspan = array();

    private $_cssParser = null;

    /**
     * Create a new PHPExcel_Reader_HTML
     */
    public function __construct()
    {
        if (class_exists('\TijsVerkoyen\CssToInlineStyles\CssToInlineStyles'))
            $this->_cssParser = new PHPExcel_Parsers_CssParser(new \TijsVerkoyen\CssToInlineStyles\CssToInlineStyles());

        $this->_readFilter = new PHPExcel_Reader_DefaultReadFilter();
    }

    /**
     * Validate that the current file is an HTML file
     *
     * @return boolean
     */
    protected function _isValidFormat()
    {
        //	Reading 2048 bytes should be enough to validate that the format is HTML
        $data = fread($this->_fileHandle, 2048);
        if ((strpos($data, '<') !== FALSE) &&
                (strlen($data) !== strlen(strip_tags($data)))) {
            return TRUE;
        }

        return FALSE;
    }

    /**
     * Loads PHPExcel from file
     *
     * @param  string                    $pFilename
     * @return PHPExcel
     * @throws PHPExcel_Reader_Exception
     */
    public function load($pFilename)
    {
        // Create new PHPExcel
        $objPHPExcel = new PHPExcel();

        // Load into this instance
        return $this->loadIntoExisting($pFilename, $objPHPExcel);
    }

    /**
     * Set input encoding
     *
     * @param string $pValue Input encoding
     */
    public function setInputEncoding($pValue = 'ANSI')
    {
        $this->_inputEncoding = $pValue;

        return $this;
    }

    /**
     * Get input encoding
     *
     * @return string
     */
    public function getInputEncoding()
    {
        return $this->_inputEncoding;
    }

    //	Data Array used for testing only, should write to PHPExcel object on completion of tests
    protected $_dataArray = array();
    protected $_tableLevel = 0;
    protected $_nestedColumn = array('A');

    protected function _setTableStartColumn($column)
    {
        if ($this->_tableLevel == 0)
            $column = 'A';
        ++$this->_tableLevel;
        $this->_nestedColumn[$this->_tableLevel] = $column;

        return $this->_nestedColumn[$this->_tableLevel];
    }

    protected function _getTableStartColumn()
    {
        return $this->_nestedColumn[$this->_tableLevel];
    }

    protected function _releaseTableStartColumn()
    {
        --$this->_tableLevel;

        return array_pop($this->_nestedColumn);
    }

    protected function _flushCell($sheet, $column, $row, &$cellContent)
    {
        if (is_string($cellContent)) {
            //	Simple String content
            if (trim($cellContent) > '') {
                //	Only actually write it if there's content in the string
//				echo 'FLUSH CELL: ' , $column , $row , ' => ' , $cellContent , '<br />';
                //	Write to worksheet to be done here...
                //	... we return the cell so we can mess about with styles more easily
                $sheet->setCellValue($column . $row, $cellContent, true);
                $this->_dataArray[$row][$column] = $cellContent;
            }
        } else {
            //	We have a Rich Text run
            //	TODO
            $this->_dataArray[$row][$column] = 'RICH TEXT: ' . $cellContent;
        }
        $cellContent = (string) '';
    }

    protected function _processDomElement(DOMNode $element, $sheet, &$row, &$column, &$cellContent, $format = null)
    {
        foreach ($element->childNodes as $child) {
            if ($child instanceof DOMText) {
                $domText = preg_replace('/\s+/u', ' ', trim($child->nodeValue));
                if (is_string($cellContent)) {
                    //	simply append the text if the cell content is a plain text string
                    $cellContent .= $domText;
                } else {
                    //	but if we have a rich text run instead, we need to append it correctly
                    //	TODO
                }
            } elseif ($child instanceof DOMElement) {
//				echo '<b>DOM ELEMENT: </b>' , strtoupper($child->nodeName) , '<br />';

                $attributeArray = array();
                foreach ($child->attributes as $attribute) {
//					echo '<b>ATTRIBUTE: </b>' , $attribute->name , ' => ' , $attribute->value , '<br />';
                    $attributeArray[$attribute->name] = $attribute->value;
                }

                switch ($child->nodeName) {
                    case 'meta' :
                        foreach ($attributeArray as $attributeName => $attributeValue) {
                            switch ($attributeName) {
                                case 'content':
                                    //	TODO
                                    //	Extract character set, so we can convert to UTF-8 if required
                                    break;
                            }
                        }
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        break;
                    case 'title' :
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        $sheet->setTitle($cellContent);
                        $cellContent = '';
                        break;
                    case 'span' :
                    case 'div' :
                    case 'font' :
                    case 'i' :
                    case 'em' :
                    case 'strong':
                    case 'b' :
//						echo 'STYLING, SPAN OR DIV<br />';
                        if ($cellContent > '')
                            $cellContent .= ' ';
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        if ($cellContent > '')
                            $cellContent .= ' ';
//						echo 'END OF STYLING, SPAN OR DIV<br />';
                        break;
                    case 'hr' :
                        $this->_flushCell($sheet, $column, $row, $cellContent);
                        ++$row;
                        if (isset($this->_formats[$child->nodeName])) {
                            $sheet->getStyle($column . $row)->applyFromArray($this->_formats[$child->nodeName]);
                        } else {
                            $cellContent = '----------';
                            $this->_flushCell($sheet, $column, $row, $cellContent);
                        }
                        ++$row;
                    case 'br' :
                        if ($this->_tableLevel > 0) {
                            //	If we're inside a table, replace with a \n
                            $cellContent .= "\n";
                        } else {
                            //	Otherwise flush our existing content and move the row cursor on
                            $this->_flushCell($sheet, $column, $row, $cellContent);
                            ++$row;
                        }
//						echo 'HARD LINE BREAK: ' , '<br />';
                        break;
                    case 'a' :
//						echo 'START OF HYPERLINK: ' , '<br />';
                        foreach ($attributeArray as $attributeName => $attributeValue) {
                            switch ($attributeName) {
                                case 'href':
//									echo 'Link to ' , $attributeValue , '<br />';
                                    $sheet->getCell($column . $row)->getHyperlink()->setUrl($attributeValue);
                                    if (isset($this->_formats[$child->nodeName])) {
                                        $sheet->getStyle($column . $row)->applyFromArray($this->_formats[$child->nodeName]);
                                    }
                                    break;
                            }
                        }
                        $cellContent .= ' ';
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//						echo 'END OF HYPERLINK:' , '<br />';
                        break;
                    case 'span'  :
                    case 'div'   :
                    case 'font'  :
                    case 'i'     :
                    case 'em'    :
                    case 'strong':
                    case 'b'     :

                        // Add space after empty cells
                        if ( $cellContent > '' )
                            $cellContent .= ' ';

                        // Continue processing
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent, $format);

                        // Add space after empty cells
                        if ( $cellContent > '' )
                            $cellContent .= ' ';

                        // Set the styling
                        if ( isset($this->_formats[$child->nodeName]) )
                        {
                            $sheet->getStyle($column . $row)
                                  ->applyFromArray($this->_formats[$child->nodeName]);
                        }

                        break;
                    case 'h1' :
                    case 'h2' :
                    case 'h3' :
                    case 'h4' :
                    case 'h5' :
                    case 'h6' :
                    case 'ol' :
                    case 'ul' :
                    case 'p' :
                        if ($this->_tableLevel > 0) {
                            //	If we're inside a table, replace with a \n
                            $cellContent .= "\n";
//							echo 'LIST ENTRY: ' , '<br />';
                            $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//							echo 'END OF LIST ENTRY:' , '<br />';
                        } else {
                            if ($cellContent > '') {
                                $this->_flushCell($sheet, $column, $row, $cellContent);
                                $row++;
                            }
//							echo 'START OF PARAGRAPH: ' , '<br />';
                            $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//							echo 'END OF PARAGRAPH:' , '<br />';
                            $this->_flushCell($sheet, $column, $row, $cellContent);

                            if (isset($this->_formats[$child->nodeName])) {
                                $sheet->getStyle($column . $row)->applyFromArray($this->_formats[$child->nodeName]);
                            }

                            $row++;
                            $column = 'A';
                        }
                        break;
                    case 'li' :
                        if ($this->_tableLevel > 0) {
                            //	If we're inside a table, replace with a \n
                            $cellContent .= "\n";
//							echo 'LIST ENTRY: ' , '<br />';
                            $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//							echo 'END OF LIST ENTRY:' , '<br />';
                        } else {
                            if ($cellContent > '') {
                                $this->_flushCell($sheet, $column, $row, $cellContent);
                            }
                            ++$row;
//							echo 'LIST ENTRY: ' , '<br />';
                            $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//							echo 'END OF LIST ENTRY:' , '<br />';
                            $this->_flushCell($sheet, $column, $row, $cellContent);
                            $column = 'A';
                        }
                        break;

                    case 'img':
                        $this->insertImageBySrc($sheet, $column, $row, $child);
                        break;

                    case 'table' :
                        $this->_flushCell($sheet, $column, $row, $cellContent);
                        $column = $this->_setTableStartColumn($column);
//						echo 'START OF TABLE LEVEL ' , $this->_tableLevel , '<br />';
                        if ($this->_tableLevel > 1)
                            --$row;
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//						echo 'END OF TABLE LEVEL ' , $this->_tableLevel , '<br />';
                        $column = $this->_releaseTableStartColumn();
                        if ($this->_tableLevel > 1) {
                            ++$column;
                        } else {
                            ++$row;
                        }
                        break;
                    case 'thead' :
                    case 'tbody' :
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        break;
                    case 'tr' :
                        $column = $this->_getTableStartColumn();
                        $cellContent = '';
//						echo 'START OF TABLE ' , $this->_tableLevel , ' ROW<br />';
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        ++$row;
//						echo 'END OF TABLE ' , $this->_tableLevel , ' ROW<br />';
                        break;
                    case 'th' :
                    case 'td' :
//						echo 'START OF TABLE ' , $this->_tableLevel , ' CELL<br />';
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
//						echo 'END OF TABLE ' , $this->_tableLevel , ' CELL<br />';

                        while (isset($this->rowspan[$column . $row])) {
                            ++$column;
                        }

                        $this->_flushCell($sheet, $column, $row, $cellContent);

                        if (isset($attributeArray['style']) && !empty($attributeArray['style'])) {
                            $styleAry = $this->getPhpExcelStyleArray($attributeArray['style']);

                            if (!empty($styleAry)) {
                                $sheet->getStyle($column . $row)->applyFromArray($styleAry);
                            }

                            if (isset($styleAry['width']))
                            {
                                if ($styleAry['width'] == 'auto')
                                    $sheet->getColumnDimension($column)->setAutoSize(true);
                                else
                                    $sheet->getColumnDimension($column)->setWidth((int)$styleAry['width']);
                            }
                            if (isset($styleAry['height']))
                            {
                                if ($styleAry['height'] == 'auto')
                                    $sheet->getRowDimension($row)->setAutoSize(height);
                                else
                                    $sheet->getRowDimension($row)->setRowHeight((int)$styleAry['height']);
                            }
                        }

                        if (isset($attributeArray['rowspan']) && isset($attributeArray['colspan'])) {
                            //create merging rowspan and colspan
                            $columnTo = $column;
                            for ($i = 0; $i < $attributeArray['colspan'] - 1; $i++) {
                                ++$columnTo;
                            }
                            $range = $column . $row . ':' . $columnTo . ($row + $attributeArray['rowspan'] - 1);
                            foreach (\PHPExcel_Cell::extractAllCellReferencesInRange($range) as $value) {
                                $this->rowspan[$value] = true;
                            }
                            $sheet->mergeCells($range);
                            $column = $columnTo;
                        } elseif (isset($attributeArray['rowspan'])) {
                            //create merging rowspan
                            $range = $column . $row . ':' . $column . ($row + $attributeArray['rowspan'] - 1);
                            foreach (\PHPExcel_Cell::extractAllCellReferencesInRange($range) as $value) {
                                $this->rowspan[$value] = true;
                            }
                            $sheet->mergeCells($range);
                        } elseif (isset($attributeArray['colspan'])) {
                            //create merging colspan
                            $columnTo = $column;
                            for ($i = 0; $i < $attributeArray['colspan'] - 1; $i++) {
                                ++$columnTo;
                            }
                            $sheet->mergeCells($column . $row . ':' . $columnTo . $row);
                            $column = $columnTo;
                        }

                        if (isset($attributeArray['style']) && !empty($attributeArray['style']))
                            $sheet->getStyle($column . $row)->applyFromArray($styleAry);

                        ++$column;
                        break;
                    case 'body' :
                        $row = 1;
                        $column = 'A';
                        $content = '';
                        $this->_tableLevel = 0;
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                        break;
                    default:
                        $this->_processDomElement($child, $sheet, $row, $column, $cellContent);
                }
            }
        }
    }

    /**
     * Loads PHPExcel from file into PHPExcel instance
     *
     * @param  string                    $pFilename
     * @param  PHPExcel                  $objPHPExcel
     * @return PHPExcel
     * @throws PHPExcel_Reader_Exception
     */
    public function loadIntoExisting($pFilename, PHPExcel $objPHPExcel)
    {
        // Open file to validate
        $this->_openFile($pFilename);
        if (!$this->_isValidFormat()) {
            fclose($this->_fileHandle);
            throw new PHPExcel_Reader_Exception($pFilename . " is an Invalid HTML file.");
        }
        //	Close after validating
        fclose($this->_fileHandle);

        // Create new PHPExcel
        while ($objPHPExcel->getSheetCount() <= $this->_sheetIndex) {
            $objPHPExcel->createSheet();
        }
        $objPHPExcel->setActiveSheetIndex($this->_sheetIndex);

        //	Create a new DOM object
        $dom = new domDocument;
        //	Reload the HTML file into the DOM object
        $loaded = $dom->loadHTML($this->securityScanFile($pFilename));

        // apply non-inline styles
        if ($this->_cssParser)
        {
            // Let the css parser find all stylesheets
            $this->_cssParser->findStyleSheets($dom);

            // Transform the css files to inline css and replace the html
            $html = $this->_cssParser->transformCssToInlineStyles($this->securityScanFile($pFilename));

            // Re-init dom doc
            $dom = new DOMDocument;

            // Load again with css included
            $loaded = $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        }

        if ($loaded === FALSE) {
            throw new PHPExcel_Reader_Exception('Failed to load ', $pFilename, ' as a DOM Document');
        }

        //	Discard white space
        $dom->preserveWhiteSpace = false;

        $row = 0;
        $column = 'A';
        $content = '';
        $this->_processDomElement($dom, $objPHPExcel->getActiveSheet(), $row, $column, $content);

		// Return
        return $objPHPExcel;
    }

    /**
     * Get sheet index
     *
     * @return int
     */
    public function getSheetIndex()
    {
        return $this->_sheetIndex;
    }

    /**
     * Set sheet index
     *
     * @param  int                  $pValue Sheet index
     * @return PHPExcel_Reader_HTML
     */
    public function setSheetIndex($pValue = 0)
    {
        $this->_sheetIndex = $pValue;

        return $this;
    }

	/**
	 * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
	 *
	 * @param 	string 		$xml
	 * @throws PHPExcel_Reader_Exception
	 */
	public function securityScan($xml)
	{
        $pattern = '/\\0?' . implode('\\0?', str_split('<!ENTITY')) . '\\0?/';
        if (preg_match($pattern, $xml)) {
            throw new PHPExcel_Reader_Exception('Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks');
        }
        return $xml;
    }

    /**
     * Converts an array of css style attributes to PHPExcel style values.
     * <p>
     * Any array element that is not a valid css attribute will be merged into
     * the returned array.
     *
     * @param  $mixed An array of css attributes
     * @return Array
     */
    public function getPHPExcelStyleArray()
    {
        $returnStyle = array();
        $args        = func_get_args();

        $new_args = array();
        foreach($args as $key => $css)
        {
            foreach(explode(";", $css) as $css_line)
            {
                if (trim($css_line) == "")
                    continue;
                $tmp = explode(":", $css_line);
                $new_args[$key][trim($tmp[0])] = trim(str_replace(";", '', $tmp[1]));
            }
        }
        $args = $new_args;

        foreach ($args as $key=>$css) {
        $style = array();
        // text-alignment
        if (isset($css['text-align'])) {
            $style['alignment'] = array();
            switch ($css['text-align']) {
            case 'left':$style['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
                break;
            case 'center':$style['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
                break;
            case 'right':$style['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
                break;
            }

            unset($args[$key]['text-align']);
        }

        if (isset($css['vertical-align'])) {
            $style['alignment'] = isset($style['alignment']) ? $style['alignment'] : array();

            switch ($css['vertical-align']) {
            case 'top':$style['alignment']['vertical'] = PHPExcel_Style_Alignment::VERTICAL_TOP;
                break;
            case 'middle':$style['alignment']['vertical'] = PHPExcel_Style_Alignment::VERTICAL_CENTER;
                break;
            case 'bottom':$style['alignment']['vertical'] = PHPExcel_Style_Alignment::VERTICAL_BOTTOM;
                break;
            }

            unset($args[$key]['vertical-align']);
        }

        // background-color
        if (isset($css['background-color'])) {
            $style['fill'] = array(
                'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->getColor($css['background-color'])),
            );

            unset($args[$key]['background-color']);
        }

        if (isset($css['background'])) {
            $style['fill'] = array(
                'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $this->getColor($css['background'])),
            );

            unset($args[$key]['background']);
        }

        // font-size
        if (isset($css['font-size'])) {
            $style['font']         = isset($style['font']) ? $style['font'] : array();
            $style['font']['size'] = $css['font-size'];

            unset($args[$key]['font-size']);
        }

        // font-weight
        if (isset($css['font-weight'])) {
            $style['font']         = isset($style['font']) ? $style['font'] : array();
            $style['font'][$css['font-weight']] = true;

            unset($args[$key]['font-weight']);
        }

        // font-color
        if (isset($css['color'])) {
            $style['font']          = isset($style['font']) ? $style['font'] : array();
            $style['font']['color'] = array('rgb' => $this->getColor($css['color']));

            unset($args[$key]['color']);
        }

        // border
        if (isset($css['border'])) {
            $borderParts = explode(' ', $css['border']);

            $style['borders'] = isset($style['borders']) ? $style['borders'] : array();

            $border = array(
                        'style' => $this->getBorderStyle($borderParts[0]),
                            'color' => array(
                            'rgb' => $this->getColor(end($borderParts))
                        )
                    );

            $style['borders'] = array(
                'bottom' => $border,
                'left'   => $border,
                'top'    => $border,
                'right'  => $border,
            );

            unset($args[$key]['borders']);
        }

        if (isset($css['border-top'])) {
            $borderParts = explode(' ', $css['border-top']);

            $style['borders'] = isset($style['borders']) ? $style['borders'] : array();

            $border = array(
                        'style' => $this->getBorderStyle($borderParts[0]),
                            'color' => array(
                            'rgb' => $this->getColor(end($borderParts))
                        )
                    );

            $style['borders'] = array(
                'top' => $border,
            );

            unset($args[$key]['border-top']);
        }

        if (isset($css['border-right'])) {
            $borderParts = explode(' ', $css['border-right']);

            $style['borders'] = isset($style['borders']) ? $style['borders'] : array();

            $border = array(
                        'style' => $this->getBorderStyle($borderParts[0]),
                            'color' => array(
                            'rgb' => $this->getColor(end($borderParts))
                        )
                    );

            $style['borders'] = array(
                'right' => $border,
            );

            unset($args[$key]['border-right']);
        }

        if (isset($css['border-bottom'])) {
            $borderParts = explode(' ', $css['border-bottom']);

            $style['borders'] = isset($style['borders']) ? $style['borders'] : array();

            $border = array(
                        'style' => $this->getBorderStyle($borderParts[0]),
                            'color' => array(
                            'rgb' => $this->getColor(end($borderParts))
                        )
                    );

            $style['borders'] = array(
                'bottom' => $border,
            );

            unset($args[$key]['border-bottom']);
        }

        if (isset($css['border-left'])) {
            $borderParts = explode(' ', $css['border-left']);

            $style['borders'] = isset($style['borders']) ? $style['borders'] : array();

            $border = array(
                        'style' => $this->getBorderStyle($borderParts[0]),
                            'color' => array(
                            'rgb' => $this->getColor(end($borderParts))
                        )
                    );

            $style['borders'] = array(
                'left' => $border,
            );

            unset($args[$key]['border-left']);
        }

        $returnStyle = array_merge($returnStyle, $args[$key]);
        $returnStyle = array_merge($returnStyle, $style);
        }

        return $returnStyle;
    }

    /**
     * Get the color
     * @param  string $color
     * @return string
     */
    public function getColor($color)
    {
        if (!ctype_xdigit($color))
        {
            $hex_color = @PHPExcel_Helper_HTML::colourNameLookup($color);
            if ($hex_color != null)
                $color = $hex_color;
        }

        $color = str_replace('#', '', $color);

        // If color is only 3 chars long, mirror it to 6 chars
        if ( strlen($color) == 3 )
            $color = $color . $color;

        return $color;
    }

    /**
     * Get the border style
     * @param  string $style
     * @return string
     */
    public function getBorderStyle($style)
    {
        switch ($style)
        {
            case 'solid';
                return PHPExcel_Style_Border::BORDER_THIN;
                break;

            case 'dashed':
                return PHPExcel_Style_Border::BORDER_DASHED;
                break;

            case 'dotted':
                return PHPExcel_Style_Border::BORDER_DOTTED;
                break;

            case 'medium':
                return PHPExcel_Style_Border::BORDER_MEDIUM;
                break;

			case 'thin':
                return PHPExcel_Style_Border::BORDER_THIN;
                break;

            case 'thick':
                return PHPExcel_Style_Border::BORDER_THICK;
                break;

            case 'none':
                return PHPExcel_Style_Border::BORDER_NONE;
                break;

            case 'dash-dot':
                return PHPExcel_Style_Border::BORDER_DASHDOT;
                break;

            case 'dash-dot-dot':
                return PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                break;

            case 'double':
                return PHPExcel_Style_Border::BORDER_DOUBLE;
                break;

            case 'hair':
                return PHPExcel_Style_Border::BORDER_HAIR;
                break;

            case 'medium-dash-dot':
                return PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                break;

            case 'medium-dash-dot-dot':
                return PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                break;

            case 'medium-dashed':
                return PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                break;

            case 'slant-dash-dot':
                return PHPExcel_Style_Border::BORDER_SLANTDASHDOT;
                break;

            default:
                return null;
                break;
        }
    }

    /**
     * Insert a image inside the sheet
     * @param  Worksheet $sheet
     * @param  string    $column
     * @param  integer   $row
     * @param  string    $attributes
     * @return void
     */
    protected function insertImageBySrc($sheet, $column, $row, $attributes)
    {
        // Get attributes
        $src = $attributes->getAttribute('src');
        $width = (float) $attributes->getAttribute('width');
        $height = (float) $attributes->getAttribute('height');
        $alt = $attributes->getAttribute('alt');
        $style = $attributes->getAttribute('style');

        if ($style)
            $style = $this->getPhpExcelStyleArray($style);
        else
            $style = array();

        $top = 0;
        $left = 0;
        $height = 100;
        $width = null;
        if (isset($style['top']))
            $top = (int) $style['top'];
        if (isset($style['left']))
            $left = (int) $style['left'];
        if (isset($style['height']))
            $height = (int) $style['height'];
        if (isset($style['width']))
            $width = (int) $style['width'];

        // init drawing
        $drawing = new PHPExcel_Worksheet_Drawing();

        if (filter_var($src, FILTER_VALIDATE_URL))
        {
            $headers=get_headers($src);
            if (stripos($headers[0],"200 OK") === false)
                throw new Exception('File '.$src.' not found!');
            else
            {
                $tmp_image = $this->temporaryFile('phpexcel_tmp_image_'.basename($src), file_get_contents($src));
                $src = $tmp_image;
            }
        }

        // Set image
        $drawing->setPath($src);
        $drawing->setName($alt);
        $drawing->setWorksheet($sheet);
        $drawing->setCoordinates($column . $row);
        $drawing->setResizeProportional();
        $drawing->setOffsetX($left);
        $drawing->setOffsetY($top);

        // Set height and width
        if ( $width > 0 )
            $drawing->setWidth($width);

        if ( $height > 0 )
            $drawing->setHeight($height);
    }

    private function temporaryFile($name, $content)
    {
        $file = DIRECTORY_SEPARATOR .
                trim(sys_get_temp_dir(), DIRECTORY_SEPARATOR) .
                DIRECTORY_SEPARATOR .
                ltrim($name, DIRECTORY_SEPARATOR);

        file_put_contents($file, $content);

        register_shutdown_function(function() use($file) {
            unlink($file);
        });

        return $file;
    }
}
