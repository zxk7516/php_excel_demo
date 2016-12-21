<?php
require_once __DIR__ . '/PHPExcel.php';
if (!defined('PHPEXCEL_ROOT')) {
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}
class Excel2HtmlRender{

	private $filename;
	private $objPHPExcel;


	public function __construct($filename ){
		$this->filename = $filename;
		$this->objPHPExcel = \PHPExcel_IOFactory::load( $this->filename );
	}


	public function render2( $options ){
		$skipCell = array();

		$objWorksheet = $this->objPHPExcel->getActiveSheet();

		$mergedCells = $objWorksheet->getMergeCells();

		$col_widths = array();
		foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
			$cellIterator = $row->getCellIterator();
			try {
				$cellIterator->setIterateOnlyExistingCells(true);
					// This loops through all cells,
					//    even if a cell value is not set.
					// By default, only cells that have a value
					//    set will be iterated.
			} catch (\PHPExcel_Exception $e) {
			}
			foreach ($cellIterator as $colIdxName=>$cell) {
				$colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
				$col_widths[$colIdx] = intval( $objWorksheet->getColumnDimension($colIdxName)->getWidth() );
			}
			break;
		}
		$col_width_sum = array_sum($col_widths);

		ob_start();
		$thead = '';
		$tbody = '';
		foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
			$tmpRow = '';
			$tmpRow .= '<tr>'.PHP_EOL;
			$cellIterator = $row->getCellIterator();
			try {
				$cellIterator->setIterateOnlyExistingCells(false);
					// This loops through all cells,
					//    even if a cell value is not set.
					// By default, only cells that have a value
					//    set will be iterated.
			} catch (\PHPExcel_Exception $e) {
			}

			foreach ($cellIterator as $colIdxName=>$cell) {
				$colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
				// var_dump($colIdx);
				$rowspan = 1;
				$colspan = 1;

				if( @$skipCell[$colIdxName.$rowIdx] ){
				    echo $colIdxName.$rowIdx;
					continue;
				}
				foreach($mergedCells as $mergedCell){
					if( preg_match('/^'.preg_quote($colIdxName.$rowIdx).'\\:([a-zA-Z]+)([0-9]+)$/', $mergedCell, $matched) ){
						$maxIdxC = \PHPExcel_Cell::columnIndexFromString( $matched[1] );
						// var_dump($colIdx);
						// var_dump(\PHPExcel_Cell::stringFromColumnIndex($colIdx-1));
						// var_dump($maxIdxC);
						$maxIdxR = intval($matched[2]);
						for( $idxC=$colIdx; $idxC<=$maxIdxC; $idxC++ ){
							for( $idxR=$rowIdx; $idxR<=$maxIdxR; $idxR++ ){
								$skipCell[\PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR] = \PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR;
							}
						}
						$colspan = $maxIdxC-$colIdx+1;
						$rowspan = $maxIdxR-$rowIdx+1;
						break;
					}
				}

				// var_dump($colIdx);
				$cellTagName = 'td';
				if( $rowIdx <= $options['header_row'] || $colIdx <= $options['header_col'] ){
					$cellTagName = 'th';
				}
				$cellValue = $cell->getFormatedValue();
				switch( $options['cell_renderer'] ){
					case 'text':
						$cellValue = htmlspecialchars($cellValue);
						$cellValue = preg_replace('/\r\n|\r|\n/', '<br />', $cellValue);
						break;
					case 'html':
						break;
//					case 'markdown':
//						$cellValue = \Michelf\MarkdownExtra::defaultTransform($cellValue);
//						break;
				}
				// var_dump( $cell->getNumberFormat() );

				$styles = array();
				$cellStyle = $cell->getStyle();
				// print('<pre>');
				// // $cellStyle->getBorders()->getOutline();
				// var_dump($cellStyle->getBorders()->getLeft()->getColor()->getRGB());
				// print('</pre>');

				if( $options['render_cell_align'] ){
					if( $cellStyle->getAlignment()->getHorizontal() != 'general' ){
						array_push( $styles, 'text-align: '.strtolower($cellStyle->getAlignment()->getHorizontal()).';' );
					}
				}
				if( $options['render_cell_height'] ){
					array_push( $styles, 'height: '.intval($objWorksheet->getRowDimension($rowIdx)->getRowHeight()).'px;' );
				}
				if( $options['render_cell_font'] ){
					array_push( $styles, 'color: #'.strtolower($cellStyle->getFont()->getColor()->getRGB()).';' );
					array_push( $styles, 'font-weight: '.($cellStyle->getFont()->getBold()?'bold':'normal').';' );
					array_push( $styles, 'font-size: '.intval($cellStyle->getFont()->getsize()/12*100).'%;' );
				}
				if( $options['render_cell_background'] ){
					array_push( $styles, 'background-color: #'.strtolower($cellStyle->getFill()->getStartColor()->getRGB()).';' );
				}
				if( $options['render_cell_vertical_align'] ){
					$verticalAlign = strtolower($cellStyle->getAlignment()->getVertical());
					array_push( $styles, 'vertical-align: '.($verticalAlign=='center'?'middle':$verticalAlign).';' );
				}
				if( $options['render_cell_borders'] ){
					array_push( $styles, 'border-top: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getTop()).';' );
					array_push( $styles, 'border-right: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getRight()).';' );
					array_push( $styles, 'border-bottom: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getBottom()).';' );
					array_push( $styles, 'border-left: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getLeft()).';' );
				}


				$tmpRow .= '<'.$cellTagName.($rowspan>1?' rowspan="'.$rowspan.'"':'').($colspan>1?' colspan="'.$colspan.'"':'').''.(count($styles)?' style="'.htmlspecialchars(implode(' ',$styles)).'"':'').'>';
				$tmpRow .= $cellValue;
				$tmpRow .= '</'.$cellTagName.'>'.PHP_EOL;
			}
			$tmpRow .= '</tr>'.PHP_EOL;

			if( $rowIdx <= $options['header_row'] ){
				$thead .= $tmpRow;
			}else{
				$tbody .= $tmpRow;
			}
		}

		if( !@$options['strip_table_tag'] ){
			print '<table>'.PHP_EOL;
		}
		if( $options['render_cell_width'] ){
			print '<colgroup>'.PHP_EOL;
			foreach( $col_widths as $colIdx=>$colRow ){
				print '<col style="width:'.floatval($colRow/$col_width_sum*100).'%;" />'.PHP_EOL;
			}
			print '</colgroup>'.PHP_EOL;
		}
		if( strlen($thead) ){
			print '<thead>'.PHP_EOL;
			print $thead;
			print '</thead>'.PHP_EOL;
		}
		print '<tbody>'.PHP_EOL;
		print $tbody;
		print '</tbody>'.PHP_EOL;

		if( !@$options['strip_table_tag'] ){
			print '</table>'.PHP_EOL;
		}
		$rtn = ob_get_clean();
		return $rtn;
	} // render()

    public function render($options){
        $objWorksheets = $this->objPHPExcel->getAllSheets();
        ob_start();
        $objWorksheetNames =  $this->objPHPExcel->getSheetNames() ;
        foreach ($objWorksheets as $k => $objWorksheet){
            $skipCell = array();
            echo '<h2>'.$objWorksheetNames[$k].'</h2>';
            $mergedCells = $objWorksheet->getMergeCells();

            $col_widths = array();
            foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
                $cellIterator = $row->getCellIterator();
                try {
                    $cellIterator->setIterateOnlyExistingCells(false);
                } catch (\PHPExcel_Exception $e) {
                    var_dump($e);
                }
                foreach ($cellIterator as $colIdxName=>$cell) {
                    $colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
                    $col_widths[$colIdx] = intval( $objWorksheet->getColumnDimension($colIdxName)->getWidth() );
                }
                break;
            }
            $col_width_sum = array_sum($col_widths);
            $thead = '';
            $tbody = '';
            foreach ($objWorksheet->getRowIterator() as $rowIdx=>$row) {
                $tmpRow = '';
                $tmpRow .= '<tr>'.PHP_EOL;
                $cellIterator = $row->getCellIterator();
                try {
                    $cellIterator->setIterateOnlyExistingCells(false);
                    // This loops through all cells,
                    //    even if a cell value is not set.
                    // By default, only cells that have a value
                    //    set will be iterated.
                } catch (\PHPExcel_Exception $e) {
                }
                foreach ($cellIterator as $colIdxName=>$cell) {
                    $colIdx = \PHPExcel_Cell::columnIndexFromString( $colIdxName );
                    // var_dump($colIdx);
                    $rowspan = 1;
                    $colspan = 1;

                    if( @$skipCell[$colIdxName.$rowIdx] ){
                        continue;
                    }
                    foreach($mergedCells as $mergedCell){
                        if( preg_match('/^'.preg_quote($colIdxName.$rowIdx).'\\:([a-zA-Z]+)([0-9]+)$/', $mergedCell, $matched) ){
                            $maxIdxC = \PHPExcel_Cell::columnIndexFromString( $matched[1] );
                            // var_dump($colIdx);
                            // var_dump(\PHPExcel_Cell::stringFromColumnIndex($colIdx-1));
                            // var_dump($maxIdxC);
                            $maxIdxR = intval($matched[2]);
                            for( $idxC=$colIdx; $idxC<=$maxIdxC; $idxC++ ){
                                for( $idxR=$rowIdx; $idxR<=$maxIdxR; $idxR++ ){
                                    $skipCell[\PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR] = \PHPExcel_Cell::stringFromColumnIndex($idxC-1).$idxR;
                                }
                            }
                            $colspan = $maxIdxC-$colIdx+1;
                            $rowspan = $maxIdxR-$rowIdx+1;
                            break;
                        }
                    }

                    // var_dump($colIdx);
                    $cellTagName = 'td';
                    if( $rowIdx <= $options['header_row'] || $colIdx <= $options['header_col'] ){
                        $cellTagName = 'th';
                    }
                    $cellValue = $cell->getFormattedValue();
                    switch( $options['cell_renderer'] ){
                        case 'text':
                            $cellValue = htmlspecialchars($cellValue);
                            $cellValue = preg_replace('/\r\n|\r|\n/', '<br />', $cellValue);
                            break;
                        case 'html':
                            break;
                        case 'markdown':
                            $cellValue = \Michelf\MarkdownExtra::defaultTransform($cellValue);
                            break;
                    }
                    // var_dump( $cell->getNumberFormat() );

                    $styles = array();
                    $cellStyle = $cell->getStyle();
                    // print('<pre>');
                    // // $cellStyle->getBorders()->getOutline();
                    // var_dump($cellStyle->getBorders()->getLeft()->getColor()->getRGB());
                    // print('</pre>');

                    if( $options['render_cell_align'] ){
                        if( $cellStyle->getAlignment()->getHorizontal() != 'general' ){
                            array_push( $styles, 'text-align: '.strtolower($cellStyle->getAlignment()->getHorizontal()).';' );
                        }
                    }
                    if( $options['render_cell_height'] ){
                        array_push( $styles, 'height: '.intval($objWorksheet->getRowDimension($rowIdx)->getRowHeight()).'px;' );
                    }
                    if( $options['render_cell_font'] ){
                        array_push( $styles, 'color: #'.strtolower($cellStyle->getFont()->getColor()->getRGB()).';' );
                        array_push( $styles, 'font-weight: '.($cellStyle->getFont()->getBold()?'bold':'normal').';' );
                        array_push( $styles, 'font-size: '.intval($cellStyle->getFont()->getsize()/12*100).'%;' );
                    }
                    if( $options['render_cell_background'] ){
                        array_push( $styles, 'background-color: #'.strtolower($cellStyle->getFill()->getStartColor()->getRGB()).';' );
                    }
                    if( $options['render_cell_vertical_align'] ){
                        $verticalAlign = strtolower($cellStyle->getAlignment()->getVertical());
                        array_push( $styles, 'vertical-align: '.($verticalAlign=='center'?'middle':$verticalAlign).';' );
                    }
                    if( $options['render_cell_borders'] ){
                        array_push( $styles, 'border-top: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getTop()).';' );
                        array_push( $styles, 'border-right: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getRight()).';' );
                        array_push( $styles, 'border-bottom: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getBottom()).';' );
                        array_push( $styles, 'border-left: '.$this->get_borderstyle_by_border($cellStyle->getBorders()->getLeft()).';' );
                    }


                    $tmpRow .= '<'.$cellTagName.($rowspan>1?' rowspan="'.$rowspan.'"':'').($colspan>1?' colspan="'.$colspan.'"':'').''.(count($styles)?' style="'.htmlspecialchars(implode(' ',$styles)).'"':'').'>';
                    $tmpRow .= $cellValue;
                    $tmpRow .= '</'.$cellTagName.'>'.PHP_EOL;
                }
                $tmpRow .= '</tr>'.PHP_EOL;

                if( $rowIdx <= $options['header_row'] ){
                    $thead .= $tmpRow;
                }else{
                    $tbody .= $tmpRow;
                }
            }
            echo '<table>';
            if( $options['render_cell_width'] ){
                echo '<colgroup>';
                foreach( $col_widths as $colIdx=>$colRow ){
                    echo '<col style="width:'.floatval($colRow/$col_width_sum*100).'%;" />';
                }
                echo '</colgroup>';
            }
            if( strlen($thead) ){
                print '<thead>';
                print $thead;
                print '</thead>';
            }
            echo '<tbody>';
            echo $tbody;
            echo '</tbody></table>';
        }
        $rtn = ob_get_clean();
        return $rtn;
    }

	private function get_borderstyle_by_border( $border ){
		$style = $border->getBorderStyle();
		$border_width = '1px';
		$border_style = 'solid';
		switch( $style ){
			case 'none':
				$border_width = '0';
				$border_style = 'none';
				break;
			case 'dashDot':
				$border_style = 'dashed';
				break;
			case 'dashDotDot':
				$border_style = 'dashed';
				break;
			case 'dashed':
				$border_style = 'dashed';
				break;
			case 'dotted':
				$border_style = 'dotted';
				break;
			case 'double':
				$border_width = '3px';
				$border_style = 'double';
				break;
			case 'hair':
				break;
			case 'medium':
				$border_width = '3px';
				break;
			case 'mediumDashDot':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'mediumDashDotDot':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'mediumDashed':
				$border_width = '3px';
				$border_style = 'dashed';
				break;
			case 'slantDashDot':
				$border_width = '3px';
				$border_style = 'solid';
				break;
			case 'thick':
				$border_width = '5px';
				$border_style = 'solid';
				break;
			case 'thin':
				$border_width = '1px';
				$border_style = 'solid';
				break;
		}
		$rtn = $border_width.' '.$border_style.' #'.strtolower($border->getColor()->getRGB()).'';
		return $rtn;
	}

    public static function optimize_options( $options ){
        $rtn = array();

        $rtn['renderer'] = @$options['renderer'].'';
        if(!strlen($rtn['renderer'])){
            $rtn['renderer'] = 'strict';
        }
        $rtn['cell_renderer'] = @$options['cell_renderer'];
        if(!strlen($rtn['cell_renderer'])){
            $rtn['cell_renderer'] = 'text';
        }
        $rtn['header_row'] = @intval($options['header_row']);
        $rtn['header_col'] = @intval($options['header_col']);
        $rtn['strip_table_tag'] = @(bool) $options['strip_table_tag'];

        // レンダーオプションの初期値を設定
        $rtn['render_cell_width']          = true;
        $rtn['render_cell_height']         = false;
        $rtn['render_cell_background']     = false;
        $rtn['render_cell_font']           = false;
        $rtn['render_cell_borders']        = false;
        $rtn['render_cell_align']          = true;
        $rtn['render_cell_vertical_align'] = false;

        if( $rtn['renderer'] == 'strict' ){
            $rtn['render_cell_width']          = true;
            $rtn['render_cell_height']         = true;
            $rtn['render_cell_background']     = true;
            $rtn['render_cell_font']           = true;
            $rtn['render_cell_borders']        = true;
            $rtn['render_cell_align']          = true;
            $rtn['render_cell_vertical_align'] = true;
        }

        if( !is_null( @$options['render_cell_width'] ) ){
            $rtn['render_cell_width'] = @(bool) $options['render_cell_width'];
        }
        if( !is_null( @$options['render_cell_height'] ) ){
            $rtn['render_cell_height'] = @(bool) $options['render_cell_height'];
        }
        if( !is_null( @$options['render_cell_background'] ) ){
            $rtn['render_cell_background'] = @(bool) $options['render_cell_background'];
        }
        if( !is_null( @$options['render_cell_font'] ) ){
            $rtn['render_cell_font'] = @(bool) $options['render_cell_font'];
        }
        if( !is_null( @$options['render_cell_borders'] ) ){
            $rtn['render_cell_borders'] = @(bool) $options['render_cell_borders'];
        }
        if( !is_null( @$options['render_cell_align'] ) ){
            $rtn['render_cell_align'] = @(bool) $options['render_cell_align'];
        }
        if( !is_null( @$options['render_cell_vertical_align'] ) ){
            $rtn['render_cell_vertical_align'] = @(bool) $options['render_cell_vertical_align'];
        }

        return $rtn;
    }
}
