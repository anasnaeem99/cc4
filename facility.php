<!--

   - - - VIEW VERSION 2.2 - - -

 * This file is part of VIEw oNLY structure.
 *
 * (c) CYBER ICT TECHNOLOGIES (PRIVATE LIMATED) <admin@cybersoftechnologies.com>
 *
 * For the full copyright and license information, please view
 * the LICENSE file at https://cybersoftechnologies.com


      _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
     ..          ..       ..     ..     ........     ..           ..           ....       ..  ..     ..          ..    ..     ..
    ..            ..     ..      ..     ..            ..         ..           ......      ... ..     ..           ..  ..       ..
   ..              ..   ..       ..     ........       ..   .   ..           ...  ...     ......     ..             ...         ..
    ..              .. ..        ..     ..              .. ... ..             ......      .. ...     ..             ..         ..
     ..              ...         ..     ........         ... ...               ....       ..  ..     .......       ..         ..
       -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-


    ****    * *     *     ...     .    ..  .  .    .    .
    *****    *            .  .   ...   . . .  .   ...   .
    ****    *       *     ...   .   .  .  ..  .  .   .  ....

-->

<?php
$myAction = '';
if(explode('/', $_SERVER['REQUEST_URI'])[4] == $this->global['pageTitle'][0]->url){
  $myAction = explode('/', $_SERVER['REQUEST_URI'])[5];
}
if(explode('/', $_SERVER['REQUEST_URI'])[3] == $this->global['pageTitle'][0]->url){
  $myAction = explode('/', $_SERVER['REQUEST_URI'])[4];
}

/* 
** .......................................................................................................................
**  $variable = $this->myGet('TABLE_NAME', 'SELECT STATEMENT', 'WHERE CLAUSE', 'ORDER_BY_COLUMN', 'ASC/DESC');
** .......................................................................................................................
*/
  require_once 'library/libraryexcel/vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    use PhpOffice\PhpSpreadsheet\Style\Color;
    use PhpOffice\PhpSpreadsheet\Style\Alignment;
    use PhpOffice\PhpSpreadsheet\Style\Border;
    use PhpOffice\PhpSpreadsheet\Cell\DataValidation;

    function setlookup($cell, $selectSheet, $assignCell, $sheetName){
      $arr = array();
      for ($i=2; $i < 501; $i++){
        $myCell = $cell . $i;
        $formula = '=IFERROR(VLOOKUP('.$myCell.', '.$selectSheet.'!A:B, 2, FALSE), "")';
        array_push($arr, $formula);
      }
      $chunck = array_chunk($arr, 1);
      $sheetName->fromArray($chunck, NULL, $assignCell);
    }
    
    function prpu($cell, $assignCell, $sheetName){

      $arr = array();
      for ($i=2; $i < 501; $i++){
        $myCell = $cell . $i;
        $formula = '=IF(MID('.$myCell.',9,2)="PR",SupplyChainRiskLevel!$H$2,IF(MID('.$myCell.',10,2)="PR",SupplyChainRiskLevel!$H$2,IF(MID('.$myCell.',9,2)="PU",SupplyChainRiskLevel!$G$2,IF(MID('.$myCell.',10,2)="PU",SupplyChainRiskLevel!$G$2,IF(MID('.$myCell.',8,2)=" U",100,"")))))';
        array_push($arr, $formula);
      }
      $chunck = array_chunk($arr, 1);
      $sheetName->fromArray($chunck, NULL, $assignCell);

    }

      // STANDARD FUNCTION USE FOUR RISK SCORE EXCEL TEMPLATE //
    function setDataToField($sheet, $value, $column, $rangeStart, $rangeEnd, $isBold, $fontSize, $isFill, $cellBackground, $textColor){
        if($value){
        $sheet->setCellValue($column, $value);
    }   
        $sheet->getStyle($rangeStart . ':' . $rangeEnd)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        $style = $sheet->getStyle($rangeStart . ':' . $rangeEnd);
        $alignment = $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $font = $style->getFont();
        $font->setSize($fontSize);
        $font->setBold($isBold);
        $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color($textColor));

        if($isFill){
            $fill = $style->getFill();
            $fill->setFillType(Fill::FILL_SOLID);
            $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color($cellBackground));
        }
      }
  //
// Get Queries

// Lookup Query
$language = $_SESSION['myLanguage'];
$user = $this->myGet('tbls_user', 'BaseTbl.*', 'BaseTbl.status = "Active"', 'BaseTbl.id', 'desc');
$region = $this->myGet('tbl_region', 'BaseTbl.*', '(SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) = "'.$this->countryName.'"', 'BaseTbl.id', 'asc');
$city = $this->myGet('tbl_city', 'BaseTbl.*', '(SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) = "'.$this->countryName.'" AND BaseTbl.cityName IS NOT NULL', 'BaseTbl.id', 'asc');
//.. facility data based on php memory limit ..//
$facility = $this->myGet('tbl_supplychainrisklevel', 'BaseTbl.*', 'BaseTbl.status = "Active"', 'BaseTbl.id', 'asc');
$facilityRiskLevel = $this->myGet('tbl_supplychainrisklevel', 'BaseTbl.*', 'BaseTbl.status = "Active"', 'BaseTbl.id', 'asc');

// Edit / View Query
$recordsEdit = $this->myGet('tbl_facility', 'BaseTbl.*', 'BaseTbl.id ="'.$this->formId.'"', 'BaseTbl.id', 'asc');

if($myAction == 'edit'){
  if($this->roleId <> 3){
    $_SESSION['session']->error = 'You are not authorized to edit this record.';
    header("Location:".$this->base_url().'/Home/'.'VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup');exit();
  }
}
$whereClause = '1 AND BaseTbl.module = "Medicine" ';
if (isset ($_POST[ 'Submit' ])){

    if(!empty($_POST['id'])){
      $whereClause .= ' AND BaseTbl.id = "'.$_POST['id'].'"';
    }

    if(!empty($_POST['regionId'])){
      $whereClause .= ' AND (SELECT tbl_region.id FROM tbl_region LEFT JOIN tbl_city ON tbl_city.regionId = tbl_region.id WHERE tbl_city.id = BaseTbl.cityId) = "'.$_POST['regionId'].'"';
    }

    if(!empty($_POST['cityId'])){
      $whereClause .= ' AND BaseTbl.cityId = "'.$_POST['cityId'].'"';
    }

    if(!empty($_POST['facilityName'])){
      $whereClause .= ' AND BaseTbl.facilityName LIKE "%'.$_POST['facilityName'].'%"';
    }

    if(!empty($uploadDateFrom)){
        $whereClause .= ' AND BaseTbl.uploadDate >= "'.$uploadDateFrom.'"';
      }

    if(!empty($uploadDateTo)){
      $whereClause .= ' AND BaseTbl.uploadDate <= "'.$uploadDateTo.'"';
    }

    if(!empty($_POST['facilityRiskLevelId'])){
      $whereClause .= ' AND BaseTbl.facilityRiskLevelId IN ('.$_POST['facilityRiskLevelId'].')';
    }

    if(!empty($_POST['rows'])){
      $limit .= ' '.$_POST['rows'].' ';
    }

    $pageId = '';
    if(!empty($_POST['pagination'])){
    $whereClause .= ' AND BaseTbl.id >= '.$_POST['pagination'];
    }

    if(!empty($_POST['rows'])){
      if($_POST['rows'] != 'All'){
        $records = $this->myGet('tbl_facility', 'BaseTbl.*, (SELECT tbls_user.roleId FROM tbls_user WHERE tbls_user.id = BaseTbl.updatedby) as roleId, (SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) as countryName', $whereClause.' AND (SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) = "'.$this->countryName.'"', 'BaseTbl.id', 'desc limit '.$limit.' ');
      }
    }

    if($_POST['rows'] == 'All'){
      ob_clean();
      $records = $this->myGet('tbl_facility', 'BaseTbl.*, (SELECT tbls_user.roleId FROM tbls_user WHERE tbls_user.id = BaseTbl.updatedby) as roleId, (SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) as countryName', $whereClause.' AND (SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) = "'.$this->countryName.'"', 'BaseTbl.id', 'desc');

      // Create new excel file
      $spreadsheet = new Spreadsheet();
      // Add data to first worksheet
      $sheet1 = $spreadsheet->getActiveSheet();

      $count1 = 2;
      $translateMe = 'ID';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'A1', 'A1', 'A1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Country';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'B1', 'B1', 'B1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Region';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'C1', 'C1', 'C1', true, 12, true , 'ffffff','000000');
      $translateMe = 'City / District / Town Name';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'D1', 'D1', 'D1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Facility Name';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'E1', 'E1', 'E1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Facility Type';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'F1', 'F1', 'F1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Supply Chain Weight';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'G1', 'G1', 'G1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Address';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'H1', 'H1', 'H1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Latitude';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'I1', 'I1', 'I1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Longitude';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'J1', 'J1', 'J1', true, 12, true , 'ffffff','000000');
      $translateMe = 'Google Map URL';
      setDataToField($sheet1, (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'K1', 'K1', 'K1', true, 12, true , 'ffffff','000000');

      foreach($records as $record){
        foreach($region as $detail){
          foreach($city as $detail1){  
            if($detail->id == $detail1->regionId){
              if($record->cityId == $detail1->id){
        $sheet1->setCellValue("A" . $count1, $record->id);
        $sheet1->setCellValue("B" . $count1, $record->countryName);
        $sheet1->setCellValue("C" . $count1, $detail->regionName);
        $sheet1->setCellValue("D" . $count1, $detail1->cityName);
        $sheet1->setCellValue("E" . $count1, $record->facilityName);
        foreach ($facilityRiskLevel as $detail2){
          if($detail2->id == $record->facilityRiskLevelId){
        $sheet1->setCellValue("F" . $count1, $detail2->facilityType.'  '.$detail2->facilityChainCategory);
        $sheet1->setCellValue("G" . $count1, $record->supplyChainWeight);
        $sheet1->setCellValue("H" . $count1, $record->address);
        $sheet1->setCellValue("I" . $count1, $record->latitude);
        $sheet1->setCellValue("J" . $count1, $record->longitude);
        $sheet1->setCellValue("K" . $count1, $record->googleMapURL);
        $count1++;
      }}}}}}}

      $spreadsheet->setActiveSheetIndex(0);
    // Redirect output to a client’s web browser (Excel2007)
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Facility.xlsx"');
    header('Content-Type: text/html; charset=utf-8');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');
    // If you're serving to IE over SSL, then the following may be needed
    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
    header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header ('Pragma: public'); // HTTP/1.0
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;



    }

  }
  
  $facilities = $this->myGet('tbl_facility', 'COUNT(*) as resultCount', $whereClause.' AND BaseTbl.module = "Medicine" AND ((SELECT tbls_country.countryName FROM tbls_country LEFT JOIN tbls_user ON tbls_country.id = tbls_user.countryId WHERE tbls_user.id = BaseTbl.updatedby) = "'.$this->countryName.'" OR BaseTbl.updatedby IS NULL)', 'BaseTbl.id', 'desc');

  $pagination = @$facilities[0]->resultCount/1000;
  $l = @$facilities[0]->resultCount;
  $latestId = $records[0]->id;

  for ($i=0; $i < $pagination; $i++){
    $h = $l-1000;
    if($h < 0){
      $h = 0;
    }
    $myArray1[$latestId] = $l.' - '.$h;
    $latestId = $latestId - 1000;
    $l=$l-1000;
  }


  if (isset ($_POST[ 'deleteDate' ])){
    
    $data['isDeleted'] = 1;
    $result = $this->recordAjaxUpdate('updatedby = "'.$this->userId.'" AND uploadDate', $_POST['deleteDate'], $data, 'tbl_facility');
    header("Location:".$this->base_url().'/Home/'.'VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup');exit();

  }



// Uncomment $apiKey if you want someone outsider to access this module using api.
//$apiKey = 'KeyfR1DL2MTtlgCIz29NNo1pgHAPI7UO.';

// Colors = ['warning', 'info', 'primary', 'success', 'danger', 'secondary', 'default', 'lime', 'maroon', 'brown', 'pink', 'purple', 'yellow', 'black'];
$stages = ['Active' => 'success', 'Inactive' => 'warning'];
?>
<head>
  
</head>

<!-- Content Wrapper. Contains page content -->
<div class="content-wrapper">
  <!-- Content Header (Page header) -->
  <section class="content-header">
    <div class="container-fluid">
      <div class="row mb-2">
        <div class="col-sm-6">
          <h1><i <?php echo "class='".$this->global['pageTitle'][0]->icon."'"; ?>></i> <span class="translateMe"><?php echo $this->global['pageTitle'][0]->friendlyName; ?></span> <a href="#" data-toggle="modal" data-target="#riskDescription" data-backdrop="static" class="btn btn-warning" ><span class="translateMe">Need Help</span> <i class="fa fa-question-circle"></i></a></h1>
        </div>
        <div class="col-sm-6">
          <h1><a href="<?php if($myAction == 'lookup'){echo $this->base_url().'/';}else{echo $this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup';} ?>" style="background-color: #3e8193 !important;" <?php if($myAction == 'edit'){echo 'onclick="return confirm(\'Changes may not be saved.\')"';} ?> class="btn btn-secondary float-right"><i class="fa fa-arrow-left"></i> <span class="translateMe">Go Back</span></a></h1>
        </div>
      </div>
    </div><!-- /.container-fluid -->
  </section>

  <!-- Main content -->
  <section class="content">
    <div class="row">
      <div class="col-md-8">
      </div>
      <div class="col-md-4 float-right">
        <?php
            if($_SESSION['session']->error)
            {
        ?>
        <div class="alert alert-danger alert-dismissible">
            <button type="button" class="close" data-dismiss="alert" aria-hidden="true">&times;</button>
            <h4><i class="icon fa fa-warning"></i> <span class="translateMe">Alert!</span></h4>
            <span class="translateMe"><?php echo $_SESSION['session']->error; $_SESSION['session']->error = ''; ?></span>
        </div>
        <?php } ?>
        <?php  
            if($_SESSION['session']->success)
            {
        ?>
        <div class="alert alert-success alert-dismissable">
            <button type="button" class="close" data-dismiss="alert" aria-hidden="true">&times;</button>
            <h4><i class="icon fa fa-check"></i> <span class="translateMe">Alert!</span></h4>
            <span class="translateMe"><?php echo $_SESSION['session']->success; $_SESSION['session']->success = ''; ?></span>
        </div>
        <?php } ?>
        
        <div class="row">
            <div class="col-md-12">

            </div>
        </div>
      </div>
    </div>
     <?php if($myAction == 'lookup'){ ?>
    <div class="container-fluid">
      <?php echo '<form id="myForm" action="'.$this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup" enctype="multipart/form-data" method="post" accept-charset="utf-8" target="_blank">'; ?>
        <div class="card card-success card-outline1">
          <div class="card-header" data-toggle="collapse" data-target=".searchFilter">
            <h3 class="card-title"><i class="fa fa-search"></i> <span class="translateMe">Search Filters</span></h3>

          </div>
          <!-- /.card-header -->
          <div class="card-body cardBodyTransaction searchFilter collapse">
            <div class="row">

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'ID'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'id'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['id'], 'text', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Region / State / Province Name'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'regionId'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['regionId'], 'select', true, 'form-control select2', $region, ['regionName'], '', ''); ?>
                </div>
              </div>
              
              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'City / District / Town Name'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'cityId'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['cityId'], 'select', true, 'form-control select2', $city, ['cityName'], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Facility Name'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'facilityName'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['facilityName'], 'text', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div> 

              <!-- <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Facility Risk Level'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'facilityRiskLevelId'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['facilityRiskLevelId'], 'select', true, 'form-control select2', $facilityRiskLevel, ['facilityType'], '', ''); ?>
                </div>
              </div> -->

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Upload Date From'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'uploadDateFrom'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['uploadDateFrom'], 'date', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Upload Date To'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'uploadDateTo'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['uploadDateTo'], 'date', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Facility Type'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'facilityRiskLevelId'; ?>
                  <select class="form-control select2" multiple id="<?php echo @$column; ?>" name="<?php echo @$column; ?>" onchange="selectAllFacilityLevels($(this), $(this).val());">
                    <option value="">All</option>
                    <?php
                      if(!empty($facilityRiskLevel))
                      {
                        foreach ($facilityRiskLevel as $record)
                        {

                          ?>
                          <option value="<?php echo $record->id ?>"<?php $x = 0; $facilityRiskLevelId = explode(',', @$recordsEdit[0]->$column); $limit = count($facilityRiskLevelId); while($x <= $limit - 1) { if($facilityRiskLevelId[$x] == $record->id) {echo 'selected';}
                              $x++; }?>><?php echo $record->facilityType.' '.$record->facilityChainCategory; ?></option>
                          <?php
                        }
                      }
                    ?>
                  </select>
                </div>
                <input type="hidden" id="facilityLevelHidden" name="facilityRiskLevelId" value="">
              </div>
            
            <?php if($_POST['rows'] == '1000'){ ?>
             <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Pagination (Only works with 1000 records)'; ?>
                  <label><?php echo $label; ?></label>
                  <?php $column = 'pagination'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['pagination'], 'select', true, 'form-control select2', $myArray1, [], '', ''); ?>
                </div>
              </div>
            <?php } ?>  

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'No. of Records'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'rows'; ?>
                  <?php $myArray = ['10' => '10', '50' => '50', '100' => '100', '500' => '500', '1000' => '1000', 'All' => 'All ('.@$facilities[0]->resultCount. ') Excel Output']; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['rows'], 'select', true, 'form-control select2 required', $myArray, [], '', ''); ?>
                </div>
              </div>
              

            </div>
          </div>
          <!-- /.card-body -->
          <div id="go" class="card-footer text-right searchFilter collapse" style="border-top: 3px solid #28a745 !important">
            <?php echo '<button type="submit" class="btn btn-success" name="Submit"><span class="translateMe">Search</span></button>'; ?>
          </div>
        <!-- /.card -->
        </div>
        <!-- /.row -->
      <?php echo '</form>'; ?>
    </div>

    <div class="container-fluid">
      <?php echo '<form id="myForm" action="'.$this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup" enctype="multipart/form-data" method="post" accept-charset="utf-8">'; ?>
        <div class="card card-danger card-outline1" data-toggle="//collapse" data-target="//.deleteData">
          <div class="card-header" data-toggle="collapse" data-target=".deleteData">
            <h3 class="card-title"><i class="fa fa-trash"></i> <span class="translateMe">Delete Data</span></h3>

          </div>
          <!-- /.card-header -->
          <div class="card-body cardBodyTransaction deleteData collapse">
            <div class="row">

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Delete Date'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'deleteDate'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$_POST['deleteDate'], 'date', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>
              

            </div>
          </div>
          <!-- /.card-body -->
          <div id="go" class="card-footer text-right deleteData collapse" style="border-top: 3px solid #dc3545 !important">
            <?php echo '<button type="submit" class="btn btn-danger"><span class="translateMe">Delete</span></button>'; ?>
          </div>
        <!-- /.card -->
        </div>
        <!-- /.row -->
      <?php echo '</form>'; ?>
    </div>

    <div class="container-fluid">
      <div class="row">
        <div class="col-12">
          <div class="card card-primary card-outline">
            <div class="card-header">
              <h3 class="card-title"><span class="translateMe">Lookup</span></h3>
              <div class="card-tools">
                <?php
                  echo '<a href="#" data-toggle="modal" data-target="#uploaddata" data-backdrop="static" class="btn btn-secondary"><i class="fa fa-cloud-upload-alt"></i> <span class="translateMe">Upload Data</span></a>';
                ?>
                <?php if($myAction == 'lookup'){echo '<a class="btn"><form action="'.$this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/lookup" enctype="multipart/form-data" method="post" accept-charset="utf-8" target="_blank">';}?>
                  <?php if($myAction == 'lookup'){echo '<button type="submit" class="btn btn-success" name="downloadTemplate"><i class="fa fa-cloud-download-alt"></i> <span class="translateMe">Download Template</span></button>';}?>
                <?php if($myAction == 'lookup'){echo '</form></a>';}?>
                <?php if($this->roleId <> 4){
                  echo '<a href="'.$this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/add'.'" onclick="return confirm(\'Are you sure you want to add new record?\')" class="btn btn-primary"><i class="fa fa-plus"></i> <span class="translateMe">New Facility</span></a>';
                }
                ?>
                  
              </div>
            </div>
            <!-- /.card-header -->
            <div class="card-body table-responsive cardBodyTransaction">
             
              <table id="table" class="table table-bordered table-striped">
                <thead>
                <tr>
                  <th><span class="translateMe">S.#</span></th>
                  <th><span class="translateMe">ID</span></th>
                  <?php if($record->roleId <> 2){ ?>
                  <th><span class="translateMe">Country ok</span></th>
                  <?php } ?>
                  <th><span class="translateMe">Region</span></th>
                  <th><span class="translateMe">City / District / Town Name</span></th>
                  <th><span class="translateMe">Facility Name</span></th>
                  <th><span class="translateMe">Facility Type</span></th>
                  <th><span class="translateMe">Supply Chain Weight</span></th>
                  <th><span class="translateMe">Address</span></th>
                  <th><span class="translateMe">Latitude</span></th>
                  <th><span class="translateMe">Longitude</span></th>
                  <th><span class="translateMe">Google Map URL</span></th>
                  <th><span class="translateMe">Upload Date</span></th>
                  <th><span class="translateMe">Status</span></th>
                  <th class="text-center"><span class="translateMe">Action</span></th>
                </tr>
                </thead>
                <tbody>
                  <?php $sn=1; ?>
                  <?php
                  if(!empty($records))
                  {
                      foreach($records as $record)
                      {
                        $seenBy = explode(",",$record->seenBy);

                        // for($i=1; $i < 50000; $i++){ 
                        //   $data[$i]->facilityName = 'abc';
                        //   $data[$i]->masterId = '1';
                        // }

                        // $result = $this->recordAjaxBulkSave($data, 'tbl_facility');

                        foreach ($user as $detail){
                          if($detail->id == $record->updatedby){
                            if($detail->countryId <> $this->countryId){
                              continue 2;
                            }
                          }
                        }

                  ?>
                  <tr>
                    <td><?=$sn?>.</td>
                    <td><?php echo $record->id; ?></td>
                    <td>
                      <?php 
                        if($record->roleId <> 2){
                          echo $record->countryName;
                            } 
                          else{
                            echo '';
                          }
                      ?>
                    </td>
                    <td>
                      <?php 
                    foreach($region as $detail){
                      foreach($city as $detail1){  
                        if($detail->id == $detail1->regionId){
                          if($record->cityId == $detail1->id){
                          echo $detail->regionName;
                          break 2;
                        }}}}?>
                    </td>

                    <td><?php if(!empty($city)){foreach ($city as $detail){if($detail->id == $record->cityId){ echo $detail->cityName; }}}?></td>
                    <td><?php echo $record->facilityName; ?></td>
                    <td><?php if(!empty($facilityRiskLevel)){foreach ($facilityRiskLevel as $detail){if($detail->id == $record->facilityRiskLevelId){ echo $detail->facilityType.' '.$detail->facilityChainCategory; }}}?></td>
                    <td><?php echo $record->supplyChainWeight; ?></td>
                    <td><?php echo $record->address; ?></td>
                    <td><?php echo $record->latitude; ?></td>
                    <td><?php echo $record->longitude; ?></td>
                    <td><?php echo $record->googleMapURL; ?></td>
                    <td><?php echo $record->uploadDate; ?></td>
                    <td class="text-center">
                      <b><h4><span <?php foreach($stages as $stage => $color){if($record->status == $stage){echo "class='badge bg-$color'";}} ?>><?php echo $record->status; ?></span></h4></b>
                    </td>
                    <td class="text-center widthMaxContent">
                      <div class="btn-group">

                        <a href="<?php echo $this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/view/'.$record->id; ?>" class="btn btn-success"><i class="fa fa-eye"></i></a>
                        <?php if($this->countryName == $record->countryName){ ?>
                        <a href="<?php echo $this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/edit/'.$record->id; ?>" class="btn btn-primary" onclick="return confirm('Are you sure you want to edit this record?')"><i class="fa fa-pencil-alt"></i></a>
                        <?php } ?>
                      </div>
                    </td>
                  </tr>
                  <?php $sn++ ?>
                  <?php
                      }
                  }
                  ?>
                </tbody>

              </table>
            </div>
            <!-- /.card-body -->
          </div>
          <!-- /.card -->
        </div>
      </div>
      <!-- /.row -->
    </div>
    <!-- /.container-fluid -->
    <?php } ?>

    <?php if($myAction == 'edit' || $myAction == 'view'){ ?>
    <div class="container-fluid">
      <?php if($myAction == 'edit'){echo '<form id="myForm" action="'.$this->base_url().'/Home/VIEwoNLY/'.$this->global['pageTitle'][0]->url.'/submit" enctype="multipart/form-data" method="post" accept-charset="utf-8">';}?>
        <div class="card card-primary card-outline1">
          <div class="card-header">
            <h3 class="card-title"><span class="translateMe"><?php if($myAction == 'edit'){echo 'Edit';} if($myAction == 'view'){echo 'View';} ?> Details</span></h3>

          </div>
          <!-- /.card-header -->
          <div class="card-body cardBodyTransaction">
            <div class="row">

              <?php $column = 'module'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, 'Medicine', 'hidden', true, 'form-control', [], [], '', ''); ?>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'City / District / Town Name'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'cityId'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'select', true, 'form-control select2 required', $city, ['cityName'], '', ''); ?>
                </div>
              </div>
              
              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Facility Name'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'facilityName'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'text', true, 'form-control required', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Facility Risk Level'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'facilityRiskLevelId'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'select', true, 'form-control select2 required', $facilityRiskLevel, ['facilityType', ' ' , 'facilityChainCategory'], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Supply Chain Weight'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'supplyChainWeight'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'number', true, 'form-control', [], [], 'text-align: right; direction: ltr;', 'step="any" min="0"'); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Google Map URL'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'googleMapURL'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'text', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Latitude'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'latitude'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'text', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Longitude'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'longitude'; ?>
                    <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'text', true, 'form-control', [], [], '', ''); ?>
                </div>
              </div>

              <div class="col-md-12">
                <div class="form-group">
                  <?php $label = 'Address'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span></label>
                  <?php $column = 'address'; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'textarea', true, 'form-control required', [], [], '', 'rows="3"'); ?>
                </div>
              </div>

              <?php $column = 'uploadDate'; ?>
                <?php echo $this->myFormElement($label, $column, $myAction, date("Y-m-d"), 'hidden', true, 'form-control', [], [], '', ''); ?>

              <div class="col-md-6">
                <div class="form-group">
                  <?php $label = 'Status'; ?>
                  <label><?php echo $label; ?></label>
                  <?php $column = 'status'; ?>
                  <?php $myArray = ['Active' => 'Active', 'Inactive' => 'Inactive']; ?>
                  <?php echo $this->myFormElement($label, $column, $myAction, @$recordsEdit[0]->$column, 'select', true, 'form-control select2 required', $myArray, [], '', ''); ?>
                </div>
              </div>

            </div>
          </div>

          <!-- /.card-body -->
          <div class="card-footer">
            <?php if($myAction == 'edit'){echo '<button type="submit" onclick="return confirm(\'Are you sure you want to submit this record?\')" class="btn btn-primary translateMe">Submit</button>';}?>
          </div>
          <!-- /.card -->

        </div>
        <?php if($myAction == 'edit'){echo '</form>';}?>
      </div>
    <?php } ?>
  </section>
  <!-- /.content -->
</div>

<div class="modal fade" id="uploaddata">
  <div class="modal-dialog modal-lg">
    <div class="modal-content">
      <form id="myForm" action="<?php echo $this->base_url().'/Home/uploaddata/'.$this->global['pageTitle'][0]->url; ?>" enctype="multipart/form-data" method="post" accept-charset="utf-8">
        <div class="modal-header">
          <h4 class="modal-title"><span class="translateMe">Upload Data</span></h4>
          <button type="button" class="close" data-dismiss="modal" aria-label="Close">
            <span aria-hidden="true">&times;</span>
          </button>
        </div>
        <div class="modal-body" style="overflow-y: auto; max-height: 80vh;">
          <div class="container-fluid">
            <div class="row">

              <div class="col-md-12">
                <div class="form-group">
                  <?php $label = 'Attachment'; ?>
                  <label><span class="translateMe"><?php echo $label; ?></span> <i>[<font style="color: #F44336;"><span class="translateMe">File Format</span></font><font style="color: #3f51b5;"> (*.CSV)</font><font style="color: #F44336;"> <span class="translateMe">Max. File Size</span></font><font style="color: #3f51b5;"> 5 MB</font><font style="color: #F44336;"> <span class="translateMe">Max. Rows</span></font><font style="color: #3f51b5;"> 20,000</font>]</i></label>
                  <?php $column = 'attachment'; ?>
                  <div class="custom-file">
                    <input type="file" id="<?php echo @$column; ?>" name="<?php echo @$column; ?>" value="" class="custom-file-input1">
                    <!-- <label class="custom-file-label" for="<?php echo @$column; ?>">Choose file</label> -->
                  </div>
                </div>
              </div>

              <div class="col-md-12" style="display: none;">
                <!-- Upload Type = ['Save', 'Update'] -->
                <input type="hidden" name="uploadType" value="Save">
                <div class="form-group">
                  <?php $label = 'Columns'; ?>
                  <label><?php echo $label; ?></label>
                  <?php $column = 'columns'; ?>
                  <?php $myArray = ['cityName' => NULL, 'cityId' => NULL, 'facilityName' => NULL, 'facilityType' => NULL, 'facilityRiskLevelId' => NULL, 'supplyChainWeight' => NULL, 'address' => NULL, 'latitude' => NULL, 'longitude' => NULL, 'googleMapURL' => NULL, 'uploadDate' => date('Y-m-d') , 'module' => 'Medicine']; ?>
                  <?php foreach ($myArray as $key => $value){
                    echo '<input type="hidden" name="'.$column.'['.$key.']'.'" value="'.$value.'">';
                  }
                  ?>
                </div>
              </div>

              <!-- <div class="col-md-12">
                <form id="myForm" class="dropzone dz-clickable" action="<?php //echo base_url().'uploaddata/'.$pageTitle[0]->url; ?>" enctype="multipart/form-data" method="post" accept-charset="utf-8" style="border: inset #d0a41f !important;">
                <div class="dz-default dz-message"><button class="dz-button" type="button">Click to choose <b>OR</b> Drop your file here</button></div>
              </div> -->

            </div>
          </div>
        </div>
        <div class="modal-footer justify-content-between">
          <button type="submit" class="btn btn-primary" onclick="myFunction()"><span class="translateMe">Submit</span></button>
        </div>
      </form>
    </div>
  </div>
</div>
<script type="text/javascript">
function myFunction() {
  alert("Submitted Data successfully");
}
</script>

<div class="modal fade" id="riskDescription">
  <div class="modal-dialog modal-xl">
    <div class="modal-content">
      <div class="modal-header">
        <h4 class="modal-title"><span class="translateMe"><?php echo $this->global['pageTitle'][0]->friendlyName; ?></span> <span class="translateMe">Tutorial</span></h4>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close" onclick="$('#myVideo').get(0).pause();">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body" style="overflow-y: auto; max-height: 80vh;">
        <div class="container-fluid">
          <div class="row">
            <video id="myVideo" width="100%" controls>
              <source src="<?php echo $this->base_url().'/uploads/docs/tutorial_english_'.$this->global['pageTitle'][0]->url.'.mp4'; ?>" type="video/mp4">
                <?php if($language == 'english'){ ?>
                <track label="English" kind="subtitles" srclang="en" src="<?php echo $this->base_url().'/uploads/docs/tutorial_english_'.$this->global['pageTitle'][0]->url.'.vtt'; ?>" default />
                <?php } ?>
                <?php if($language == 'french'){ ?>
                <track label="French" kind="subtitles" srclang="fr" src="<?php echo $this->base_url().'/uploads/docs/tutorial_french_'.$this->global['pageTitle'][0]->url.'.vtt'; ?>" default />
                <?php } ?>
                <?php if($language == 'portuguese'){ ?>
                <track label="Potuguese" kind="subtitles" srclang="pt" src="<?php echo $this->base_url().'/uploads/docs/tutorial_potuguese_'.$this->global['pageTitle'][0]->url.'.vtt'; ?>" default />
                <?php } ?>
                <?php if($language == 'russian'){ ?> 
                <track label="Russian" kind="subtitles" srclang="ru" src="<?php echo $this->base_url().'/uploads/docs/tutorial_russian_'.$this->global['pageTitle'][0]->url.'.vtt'; ?>" default />
                <?php } ?>
            </video>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

 <script type="text/javascript">
  loadMyTable('table', true, -1);

<?php if (isset ($_POST[ 'Submit' ])){ ?>
  $('html, body').animate({
        scrollTop: $("#go").offset().top
    }, 2000);
<?php } ?>


$('#facilityRiskLevelId').change(function(){
    var str = '';
    $('#facilityRiskLevelId option:selected').each(function(index){
      str += $( this ).val();
      if (index < $('#facilityRiskLevelId option:selected').length - 1) {
        str += ',';
      }
    });
    $('#facilityLevelHidden').val(str);
});

function selectAllFacilityLevels(thisElement, value1){
    if(value1 == ''){
      $(thisElement).select2('destroy').find('option').prop('selected', 'selected').end().select2();
      var str = '';
    $('#facilityRiskLevelId option:selected').each(function(){
      if($( this ).val() != ''){
        str += $( this ).val() + ',';
      }
    });
    $('#FacilityLevelHidden').html(str);
    }
  }

</script>

<?php 
  
  if(isset($_POST['downloadTemplate'])){
    ob_clean();
    // Create new excel file
      $spreadsheet = new Spreadsheet();
      // Add data to first worksheet
      $sheet1 = $spreadsheet->getActiveSheet();
      $sheet1->setShowGridlines(false);

      // Sheet Title
      $translateMe = 'Instructions';
      $sheet1->setTitle((($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      
      // B2 Cell //
      $sheet1->mergeCells('B2:M2');
      $translateMe = 'INSTRUCTIONS TO CREATE, SAVE AND UPLOAD THE FILE';
      $sheet1->setCellValue('B2', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B2');
      $alignment = $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('c0504d'));
    //...
      
      // B4 Cell //
      $translateMe = 'Step-1';
      $sheet1->setCellValue('B4', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B4');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('000000'));
      $translateMe = 'Please Goto <Facilities> Sheet';
      $sheet1->setCellValue('C4', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
    //...

      // B5 Cell //
      $translateMe = 'Step-2';
      $sheet1->setCellValue('B5', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B5');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('000000'));
      $translateMe = 'Input all Facility related data in this sheet; remove any un-used rows before saving to csv format';
      $sheet1->setCellValue('c5', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
    //...
      
      // B6 Cell //
      $translateMe = 'Step-3';
      $sheet1->setCellValue('B6', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B6');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('000000'));
      $translateMe = 'Save <Facilities> as Simple Coma Separated Value (.CSV) as UTF-8 file formate with MS Office > = Ver. 2016 with a new name';
      $sheet1->setCellValue('c6', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      //...

      // B7 Cell //
      $translateMe = 'Step-4';
      $sheet1->setCellValue('B7', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B7');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('000000'));
      $translateMe = 'Upload newly created CSV file in the Online MedRS Tool Ver. 2.0';
      $sheet1->setCellValue('c7', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      //...
      
      // B8 Cell
      $translateMe = 'Step-5';
      $sheet1->setCellValue('B8', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B8');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('000000'));
      $translateMe = 'Verify the uploaded data in the Facility lookup using search filters.';
      $sheet1->setCellValue('c8', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      //...
      
      $translateMe = 'Notes:';
      $sheet1->setCellValue('B11', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $style = $sheet1->getStyle('B11');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('FF0000'));
      
      $translateMe = "Don't refresh or close the MedRS Tool Ver. 2.0 while uploading the offline data";
      $sheet1->setCellValue('c12', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Make sure the file you are uploading  is a CSV File and upto 5 MB';
      $sheet1->setCellValue('c13', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = "Don't change the column sequence of the CSV file";
      $sheet1->setCellValue('c14', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Make sure your data in the CSV file must start from Row No. 2';
      $sheet1->setCellValue('c15', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'You can find this settings in Excel Options ->Advance (tab) -> Use system seperator';
      $sheet1->setCellValue('c16', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Your Microsoft Office must be Genuine and Activated.';
      $sheet1->setCellValue('c17', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Your data must not contain Double Quotation mark and Backslash';
      $sheet1->setCellValue('c18', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Formulas are applied till 500 rows (If you want more you can copy as desired)';
      $sheet1->setCellValue('c19', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Please copy your data to new sheet before making that sheet csv otherwise you might see 500 blank entries because formula is pasted for 500 rows so it may consider those rows.';
      $sheet1->setCellValue('c20', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Columns with blue background will be uploaded into the system whereas the columns with white background are just for information.';
      $sheet1->setCellValue('c21', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'The fields highlighted with an Asterisk (*) are mandatory and user will be required to input required data in these highlighted columns for the smooth processing of data in the Tool.';
      $sheet1->setCellValue('c22', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));

      // Add data to second worksheet
      $sheet2 = $spreadsheet->createSheet();
      $translateMe = 'Facilities';
      $sheet2->setTitle((($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));

      $sheet2->getColumnDimension('A')->setAutoSize(true);
      $sheet2->getColumnDimension('B')->setAutoSize(true);
      $sheet2->getColumnDimension('C')->setAutoSize(true);
      $sheet2->getColumnDimension('D')->setAutoSize(true);
      $sheet2->getColumnDimension('E')->setAutoSize(true);
      $sheet2->getColumnDimension('F')->setAutoSize(true);
      $sheet2->getColumnDimension('G')->setAutoSize(true);
      $sheet2->getColumnDimension('H')->setAutoSize(true);
      $sheet2->getColumnDimension('I')->setAutoSize(true);
      $sheet2->getColumnDimension('J')->setAutoSize(true);
      
      $translateMe = 'City Name';
      $sheet2->setCellValue('A1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $dv = new DataValidation();
      $dv->setType(DataValidation::TYPE_LIST);
      $dv->setErrorStyle(DataValidation::STYLE_INFORMATION);
      $dv->setAllowBlank(true);
      $dv->setShowDropDown(true);
      $dv->setFormula1('=Cities!$A:$A');
      $sheet2->setDataValidation('A2:A500', $dv);
      $translateMe = 'City ID';
      $sheet2->setCellValue('B1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      setlookup('A', 'Cities', 'B2', $sheet2);
      
      $translateMe = 'Facility Name';
      $sheet2->setCellValue('C1', '* '.(($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Facility Type';
      $sheet2->setCellValue('D1', '* '.(($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $dv = new DataValidation();
      $dv->setType(DataValidation::TYPE_LIST);
      $dv->setErrorStyle(DataValidation::STYLE_INFORMATION);
      $dv->setAllowBlank(true);
      $dv->setShowDropDown(true);
      $dv->setFormula1('=SupplyChainRiskLevel!$A$2:$A$23');
      $sheet2->setDataValidation('D2:D500', $dv);
      $translateMe = 'Facility Type ID';
      $sheet2->setCellValue('E1',(($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe) );
      $translateMe = 'SupplyChainRiskLevel';
      setlookup('D', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe), 'E2', $sheet2);
      
      $translateMe = 'Supply Chain Weight';
      $sheet2->setCellValue('F1', '* '.(($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $dropdownValues = array('40', '60', '100');
      $dv = new DataValidation();
      $dv->setType(DataValidation::TYPE_LIST);
      $dv->setErrorStyle(DataValidation::STYLE_INFORMATION);
      $dv->setAllowBlank(true);
      $dv->setShowDropDown(true);
      $dv->setFormula1('"' . implode(',', $dropdownValues) . '"');
      $sheet2->setDataValidation('F2:F500', $dv);

      prpu('D', 'F2', $sheet2);
      
      $translateMe = 'Address';
      $sheet2->setCellValue('G1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Latitude';
      $sheet2->setCellValue('H1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Longitude';
      $sheet2->setCellValue('I1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Google Map URL';
      $sheet2->setCellValue('J1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
    //BOLD A1 //
      $sheet2->getStyle('A1:A500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('A1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
    //  

    // WITH OUT BOLD  A //
      $style = $sheet2->getStyle('A2:A500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
    // END //

   //BOLD B1 //
      $sheet2->getStyle('B1:B500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('B1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  B //
      $style = $sheet2->getStyle('B2:B500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD C1 //
      $sheet2->getStyle('C1:C500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('C1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  C //
      $style = $sheet2->getStyle('C2:C500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD D1 //
      $sheet2->getStyle('D1:D500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('D1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
    //  

    // WITH OUT BOLD  D //
      $style = $sheet2->getStyle('D2:D500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
    // END //

    //BOLD E1 //
      $sheet2->getStyle('E1:E500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('E1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  E //
      $style = $sheet2->getStyle('E2:E500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD F1 //
      $sheet2->getStyle('F1:F500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('F1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  F //
      $style = $sheet2->getStyle('F2:F500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD G1 //
      $sheet2->getStyle('G1:G500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('G1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  G //
      $style = $sheet2->getStyle('G2:G500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD H1 //
      $sheet2->getStyle('H1:H500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('H1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  H //
      $style = $sheet2->getStyle('H2:H500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD I1 //
      $sheet2->getStyle('I1:I500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('I1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  I //
      $style = $sheet2->getStyle('I2:I500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

    //BOLD J1 //
      $sheet2->getStyle('J1:J500')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
      $style = $sheet2->getStyle('J1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    //  

    // WITH OUT BOLD  J //
      $style = $sheet2->getStyle('J2:J500');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //  

    // Add data to THIRD worksheet
      $sheet3 = $spreadsheet->createSheet();
      $sheet3->setTitle('Cities');

      $sheet3->getColumnDimension('A')->setAutoSize(true);
      $sheet3->getColumnDimension('B')->setAutoSize(true);
      
      $translateMe = 'City Name';
      $sheet3->setCellValue('A1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'City ID';
      $sheet3->setCellValue('B1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      
      $count = 2;
    
      foreach($city as $cities)
      {
        $sheet3->setCellValue("A" . $count, $cities->cityName);
        $sheet3->setCellValue("B" . $count, $cities->id);
        $count++;
      }

    // A CELL BOLD //
      $style = $sheet3->getStyle('A1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      $style = $sheet3->getStyle('B1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

    // Add data to FOURTH worksheet
      $sheet4 = $spreadsheet->createSheet();
      $translateMe = 'SupplyChainRiskLevel';
      $sheet4->setTitle((($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));

      $sheet4->getColumnDimension('A')->setAutoSize(true);
      $sheet4->getColumnDimension('B')->setAutoSize(true);
      $sheet4->getColumnDimension('C')->setAutoSize(true);
      $sheet4->getColumnDimension('D')->setAutoSize(true);
      $sheet4->getColumnDimension('E')->setAutoSize(true);
      $sheet4->getColumnDimension('F')->setAutoSize(true);
      
      $translateMe = 'Compare';
      $sheet4->setCellValue('A1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Facility Type ID';
      $sheet4->setCellValue('B1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Facility Type';
      $sheet4->setCellValue('C1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Facility Supply Chain Category';
      $sheet4->setCellValue('D1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Supply Chain Sector';
      $sheet4->setCellValue('E1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Assigned Risk Scale';
      $sheet4->setCellValue('F1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Public';
      $sheet4->setCellValue('G1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));
      $translateMe = 'Private';
      $sheet4->setCellValue('H1', (($this->$language[$translateMe])? $this->$language[$translateMe]: $translateMe));

      $count = 2;
    
      foreach($facility as $record)
      { 
        $sheet4->setCellValue("B" . $count, $record->id);
        $sheet4->setCellValue("C" . $count, $record->facilityType);
        $sheet4->setCellValue("D" . $count, $record->facilityChainCategory);
        $sheet4->setCellValue("E" . $count, $record->supplyChainType);
        $sheet4->setCellValue("F" . $count, $record->supplyChainRiskScore);
        $sheet4->setCellValue("G2" , '40');
        $sheet4->setCellValue("H2" , '60');
        $count++;
      }

    // A CELL BOLD //
    for($i = 2; $i <= 60; $i++){
      $sheet4->setCellValue('A'.$i, '=CONCATENATE(C'.$i.', " ", D'.$i.')');
    }
      $style = $sheet4->getStyle('A1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);
      
      $style = $sheet4->getStyle('B1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      $style = $sheet4->getStyle('C1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      $style = $sheet4->getStyle('D1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      $style = $sheet4->getStyle('E1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      $style = $sheet4->getStyle('F1');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setBold(true);

      // WITH OUT BOLD  E //
      $style = $sheet4->getStyle('G2:G2');
      $font = $style->getFont();
      $font->setSize(12); 
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

      // WITH OUT BOLD  E //
      $style = $sheet4->getStyle('H2:H2');
      $font = $style->getFont();
      $font->setSize(12);
      $font->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK));
      $fill = $style->getFill();
      $fill->setFillType(Fill::FILL_SOLID);
      $fill->setStartColor(new \PhpOffice\PhpSpreadsheet\Style\Color('87CEEB'));
    // END //

      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $spreadsheet->setActiveSheetIndex(0);

      // Redirect output to a client’s web browser (Excel2007)
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment;filename="Facility.xlsx"');
      header('Content-Type: text/html; charset=utf-8');
      header('Cache-Control: max-age=0');
      // If you're serving to IE 9, then the following may be needed
      header('Cache-Control: max-age=1');
      
      // If you're serving to IE over SSL, then the following may be needed
      header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
      header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
      header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
      header ('Pragma: public'); // HTTP/1.0
      $writer = new Xlsx($spreadsheet);
      $writer->save('php://output');
      exit;
  }

 ?>
