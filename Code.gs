// Address to Geocode Apps Script
// By rvbautista on github
// Modified ferom ThinhDihn's work https://community.glideapps.com/t/automatic-geocoding-using-google-scripts/11760/3
// Modification includes skipping entries with existing geocodes; this reduces the requests sent to Google's Geocoder

function AddressToPosition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AddresstoGeocode");
  var cells = sheet.getRange("A6:D");
  
  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion('US');
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var address = cells.getCell(addressRow, addressColumn).getValue();
    if (address.length > 0){
      var latlong = sheet.getRange(addressRow + 5, latColumn, 1, 2);
      if (latlong.isBlank()){
    // Geocode the address and plug the lat, lng pair into the 2nd and 3rd elements of the current range row.
        location = geocoder.geocode(address);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
        if (location.status == 'OK') {
          lat = location["results"][0]["geometry"]["location"]["lat"];
          lng = location["results"][0]["geometry"]["location"]["lng"];
          latlngcomb = lat.toString() + ", " + lng.toString();
          cells.getCell(addressRow, latColumn).setValue(lat);
          cells.getCell(addressRow, lngColumn).setValue(lng);

          cells.getCell(addressRow, lngColumn+1).setValue(latlngcomb);
        }
      }
    }
  }
};
