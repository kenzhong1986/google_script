var current_sheet_name = 'sheet_1'

function getSliceData() {
  //link the sheet of valumes
  var document = SpreadsheetApp.openById('xxxxxxxxxxxx');
  var sheet = document.getSheetByName(current_sheet_name);
  var data = sheet.getDataRange().getValues();
  var volumes = []
  for (var i = 1; i < data.length; i++) {
    var item = {};
    item.name = data[i][0];
    item.totalSlices = data[i][1];
    volumes.push(item); 
  }
  console.info('volumes', volumes)
  return volumes;
}

function processData(){
  var sumOfSlices = getSliceData();
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(current_sheet_name);
  var data = sheet.getDataRange().getValues();
  var volumDetails = [];
  var current_item ;
  for (var i = 1; i < data.length; i++) {
       if(data[i][1]){
          var item = {};
          item.volume = data[i][1];
          item.date = data[i][0];
          item.quality = data[i][8];
          item.startIndex = i;
          if(current_item){
           current_item.endIndex = i -1;
         }
          let volume = sumOfSlices.find(volume => item.volume.startsWith(volume.name));
          if(volume){
            item.totalSlices = volume.totalSlices;
          }
          current_item = item;
          volumDetails.push(item);
       }else{
         if(i == data.length -1){
            current_item.endIndex = i ;
         }else{
            continue
         } 
       }
  }
volumDetails.forEach(item => {
  console.info('item ==>', item)
  sheet.getRange(item.startIndex + 1, 11).setValue(item.totalSlices);
  sheet.getRange(item.startIndex + 1, 10).setValue(item.endIndex - item.startIndex + 1);
})
console.info('valumAndIndex', volumDetails)
calculateKpi(volumDetails);
} 

function calculateKpi(volumDetails){
  //link the sheet of kpi
  var document = SpreadsheetApp.openById('xxxxxxxxxxxxxxxxxx');
  var sheet = document.getSheetByName('KPI');
  var kpiResult = [];
  
  volumDetails.forEach(item => {
    let volumeItem = kpiResult.find(kpiItem => kpiItem.date == item.date);
    if(volumeItem){
      volumeItem.total ++;
      volumeItem.quality += item.quality
    }else {
      let kpiItem = {date: item.date, total: 1, quality: item.quality};
      kpiResult.push(kpiItem);
    }

  })
  console.info('kpiResult', kpiResult)
let index = 2;
kpiResult.forEach(item => {
  sheet.getRange(index, 1).setValue(item.date);
  sheet.getRange(index, 2).setValue('QC');
  sheet.getRange(index, 3).setValue(item.total);
  sheet.getRange(index, 4).setValue(item.quality/item.total);
  index ++;
})
}
