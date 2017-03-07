<?php

/**
 * @category   ExcelHelper
 * @package    php-excel
**/


require_once ('PHPExcel/PHPExcel.php');

class Excel
{
	private $_file;
	private $_sheet;
	private $_objPHPExcel;


	/**
	 * @param string $file Excel file
	 */
	public function __construct(string $dirFile = null) {

		if ($dirFile != null) {

			$this->_file        = $dirFile;
			$this->_objPHPExcel = PHPExcel_IOFactory::load($this->_file);

	        self::getSheet();
	        self::setEncoding();

	    }

	}


	/**
	 * @param  string $file Excel File
	 */
	public function load(string $dirFile) {

		$this->_file        = $dirFile;
		$this->_objPHPExcel = PHPExcel_IOFactory::load($this->_file);

        self::getSheet();
        self::setEncoding();

	}


	/**
	 * Get first sheet
	 * @param  $this->_sheet 
	 */
	public function getSheet() {

		$this->_sheet = $this->_objPHPExcel->getSheet(0);

	}


	/**
	 * Get highest column and row
	 * @return array
	 */
	public function getInfo() {

		$aInfo = array(

			'nbColonne' => $this->_sheet->getHighestColumn(),

			'nbLigne' 	=> $this->_sheet->getHighestRow()

		);

		return $aInfo;

	}


	/**
	 * Return array of excel file
	 * @param  boolean $bEntete Spécifie si on veux intégrer l'entête dans le tableau
	 * @return array 
	 */
	public function getArrayByMat($bEntete = true) {

		$aReturn 	= array();
        $cptColonne = 1;

        foreach ($this->_sheet->getRowIterator() as $row) {

            foreach ($row->getCellIterator() as $cell) {

            	if ($bEntete) {

            		$aReturn['entete'][] = trim($cell->getValue());

            	}
            	else {

	                if ($cptColonne==1) {

	                    $matricule = trim($cell->getValue());

	                }

	                if (self::isDate($cell)) {

	                	$aReturn[$matricule][] = PHPExcel_Style_NumberFormat::toFormattedString($cell->getValue(), 'DD/MM/YYYY');

	                }
	                else {

 		                $aReturn[$matricule][] = trim($cell->getValue());

	                }

		        }
    
                $cptColonne++;
            }

            $bEntete 	= false;
            $cptColonne = 1;

        }

        return $aReturn;

	}


	/**
	 * Retourne un tableau incrémental du fichier excel
	 * @param  boolean $bEntete Spécifie si on veux intégrer l'entête dans le tableau
	 * @return array            Tableau incrémental
	 */
	public function getArrayByInc($bEntete = true) {

		$aReturn 	= array();
        $cptColonne = 1;
        $iCpt 		= 0;

        foreach ($this->_sheet->getRowIterator() as $row) {

            foreach ($row->getCellIterator() as $cell) {

            	if ($bEntete) {

            		$aReturn['entete'][] = trim($cell->getValue());

            	}
            	else {

	                if ($cptColonne==1) {

	                    $matricule = trim($cell->getValue());

	                }

	                if (self::isDate($cell)) {

	                	$aReturn[$iCpt][] = PHPExcel_Style_NumberFormat::toFormattedString($cell->getValue(), 'DD/MM/YYYY');

	                }
	                else {

		                $aReturn[$iCpt][] = trim($cell->getValue());

	                }

		        }
    
                $cptColonne++;

            }

            $bEntete 	= false;
            $cptColonne = 1;

            $iCpt++;

        }

        return $aReturn;

	}


    /**
     * Recupere le total d'une colonne
	 * @param array aData Le tableau contenant le fichier excel
	 * @param int   iCol  le numéro de la colonne a calculer.
     */
	public function getMontantTotal(array $aData, int $iCol) {

		$iMontantGlobal = 0;

		foreach ($aData as $key => $value) {

			$iMontantGlobal += $value[$iCol];

		}

		return $iMontantGlobal;

	}


	/**
	 * Ajoute un controle à l'écriture pour la fonction setCellValueByColumnAndRow
	 * @param [string] $iCol  index numérique de la colonne
	 * @param [string] $iRow  index numérique de la ligne
	 * @param [string] $aData Valeur à écrire
	 */
	public function setCellValueByCAndR($iCol, $iRow, $aData) {

		try {

	        $this->_sheet->setCellValueByColumnAndRow($iCol, $iRow, $value);
	    	return true;

	    }
	    catch(Exception $e) {

	    	return false;

	    }

	}


	/**
	 * Compare deux tableau 
	 * @param  array        $aTab1     [description]
	 * @param  array        $aTab2     [description]
	 * @param  bool|boolean $bWithInfo [description]
	 * @return [type]                  [description]
	 */
	public static function cmpArray(array $aTab1, array $aTab2, $bWithInfo = false) {

        $aReturn        = array();
        $nbDoublon      = 0;
        $nbUnique       = 0;
        $nbDifferend    = 0;
        $nbLignes       = 0;

        foreach ($aTab1 as $key => $value) {

            if (array_key_exists($key, $aTab2)) {       //Meme matricule

                if ($value[3]==$aTab2[$key][1]) {		//Meme nom

                    $aReturn['doublon'][$key] = $value;
                    $nbDoublon++;

                }
                else {                                  //Meme matricule mais nom différend.  

                    $aReturn['differend'][$key] = $value;
                    $nbDifferend++;

                }

            }
            else {                                      //Ceux présent qu'une fois

                $aReturn['unique'][$key] = $value;
                $nbUnique++;

            }

            $nbLignes++;

        }

        if ($bWithInfo) {

	        $aReturn['info']['doublon'] 	= $nbDoublon;
	        $aReturn['info']['differend'] 	= $nbDifferend;
    	    $aReturn['info']['unique'] 		= $nbUnique;
    	    $aReturn['info']['nbLigne'] 	= $nbLignes;

    	}

        return $aReturn;

	}


	/**
	 * @param   $aData array 
	 * @return 	$aReturnError error array
	 */
	public function arrayToExcel(array $aData) {

		if (empty($aData)) {

			return false;

		}

		$iCol         = 0;
		$iRow         = 1;
		$aReturnError = array();

        foreach ($aData as $uKey => $uValue) {

            foreach ($uValue as $value) {

            	$bTest = self::setCellValueByCAndR($iCol, $iRow, $value);
                
                if (!$bTest) {

                    $aReturnError[]['error'] = array($iCol, $iRow, $value);

                }
                else {

                	$aReturnError[]['succes'] = array($iCol, $iRow, $value);

                }

                $iCol++;
            }

            $iCol = 0;

            $iRow++;

        }

        return $aReturnError;

	}


	/**
	 * @param   $sOneCase string 	aManyCase array 
	 * @return 	$aReturnError error array
	 */

	/*
	public function arrayCmpToExcel(string $sOnlyOneCase = null, array $aManyCase = null) {

		$bNoParameters = false;
		$aReturnError  = array();
		$iCol          = 0;
		$iRow          = 1;

		if ($sOnlyOneCase == null && $aManyCase == null) {

			$bNoParameters = true;
			$aManyCase     = array();

		}

	    foreach ($aReturn as $aKey => $aValue) {

            switch ($aKey) {

                case 'doublon':

                   		if (!array_key_exists('doublon', $aManyCase) || $sOnlyOneCase!='doublon' || !$bNoParameters) {

                			//On ne fais rien

                		}
                		else {

                			//Penser a creer une nouvelle feuille dans la bataille pour trier tt ca
                			//Et repenser au truc pour voir si il y a pas moyen d'eviter de dupliquer.
		                    foreach ($aValue as $uKey => $uValue) {

		                        foreach ($uValue as $value) {

		                        	$bTest = self::setCellValueByCAndR($iCol, $iRow, $value);
		                            
		                            if (!$bTest)
			                            $aReturnError['doublon'][]['error'] = array($iCol, $iRow, $value);
			                        else
			                        	$aReturnError['doublon'][]['succes'] = array($iCol, $iRow, $value);

		                            $iCol++;
		                        }
		                        $iCol = 0;
		                        $iRow++;
		                    }                			
                		}
                    break;
                case 'differend':
                   		if (!array_key_exists('differend', $aManyCase) || $sOnlyOneCase!='differend' || !$bNoParameters) {
                			//On ne fais rien
                		}
                		else {
                			//Penser a creer une nouvelle feuille dans la bataille pour trier tt ca
		                    foreach ($aValue as $uKey => $uValue) {
		                        foreach ($uValue as $value) {
		                        	$bTest = self::setCellValueByCAndR($iCol, $iRow, $value);
		                            
		                            if (!$bTest)
			                            $aReturnError['differend'][]['error'] = array($iCol, $iRow, $value);
			                        else
			                        	$aReturnError['differend'][]['succes'] = array($iCol, $iRow, $value);

		                            $iCol++;
		                        }
		                        $iCol = 0;
		                        $iRow++;
		                    }                			
                		}
                    break;
                case 'unique':
	               		if (!array_key_exists('unique', $aManyCase) || $sOnlyOneCase!='unique' || !$bNoParameters) {
	            			//On ne fais rien
	            		}
	            		else {
	            			//Penser a creer une nouvelle feuille dans la bataille pour trier tt ca
		                    foreach ($aValue as $uKey => $uValue) {
		                        foreach ($uValue as $value) {
		                        	$bTest = self::setCellValueByCAndR($iCol, $iRow, $value);
		                            
		                            if (!$bTest)
			                            $aReturnError['unique'][]['error'] = array($iCol, $iRow, $value);
			                        else
			                        	$aReturnError['unique'][]['succes'] = array($iCol, $iRow, $value);

		                            $iCol++;
		                        }
		                        $iCol = 0;
		                        $iRow++;
		                    }                			
	            		}
                    break;
                default:
                    //Rien
                    break;
            }
        }

        return $aReturnError;

	}
	*/


	public function write(string $sFileName = null, $oPhpExcel = null) {

		if (is_null($sFileName)) {

			$sFileName = 'classeur1.xlsx';

		}

		if (is_null($oPhpExcel)) {

			$writer = new PHPExcel_Writer_Excel2007($oPhpExcel);

		}
        
        try {

	        $writer->save($sFileName);
	        return true;

	    }
	    catch(Exception $e) {

	    	return false;

	    }

	}


	public function isDate($cell) {

        if ($this->_objPHPExcel->getCellXfByIndex($cell->getXfIndex())->getNumberFormat()->getFormatCode() == 'mm-dd-yy') {

        	return true;

        }

        return false;

    }


    public function excelToArray() {
		
		$aReturn 	= array();
		$iCpt 		= 0;

		foreach ($this->_sheet->getRowIterator() as $row) {

			
			foreach ($row->getCellIterator() as $key => $cell) {

				if ($key == 0) {

					$aTmp = explode("/", $cell->getValue());

					$iKey = $aTmp[3];
				
				}


				$aReturn[$iKey][] = trim($cell->getValue());

            }

            $iCpt++;

        }

        return $aReturn;

    }

}

?>
