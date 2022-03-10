// import logo from './logo.svg';
import './App.css';

var XLSX = require("xlsx-js-style");

function App() {

  function fnExportToExcel(fileExtension, fileName) {
    var elt = document.getElementById("sample");
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });

    var i;
    var x = 18; // no. of columns 
    for (i = 0; i <= x; i++) {
      wb.Sheets["sheet1"][String.fromCharCode(65 + i) + "1"].s = {
        fill: {
          patternType: "solid",
          fgColor: { rgb: "FFFFAA00" }
        },
        font: {
          name: "Arial",
          sz: 14,
          bold: true,
          underline: false,
          color: { rgb: "#000000" },
        },
        alignment: {
          vertical: "center",
          horizontal: "center",
        },
      }
    };
    var a;
    var b = 10;
    for (a = 2; a <= b; a++) {
      wb.Sheets["sheet1"]["A" + a].s = {
        font: {
          name: "Calibri",
          sz: 12,
          underline: true,
          // bold: true,
          // color: { rgb: "#000000" },
        },
      }
    };

    // /* Find desired cell */
    // var desired_cell = worksheet["$A:$A"];
    // /* Get the value */
    // desired_cell.s = {
    //   font: {
    //     name: "Calibri",
    //     sz: 24,
    //     bold: true,
    //     color: { rgb: "FFFFAA00" },  
    //   },
    // };

    return XLSX.writeFile(wb, fileName + "." + fileExtension || ('MySheetName.' + (fileExtension || 'xlsx')));
  }
  return (
    <div className="App">
      <span>Enter Name: </span>
      <table id="sample" border="1" style={{ margin: 20 }}><tr><td>aircraft</td><td>workpackageName</td><td>barcode</td><td>operator</td><td>aircraftType</td><td>shortDescription</td><td>barcode</td><td>sequenceNumberPrefix</td><td>sequenceNumber</td><td>documentType</td><td>workTemplateNumber</td><td>workTemplateTitle</td><td>eventID</td><td>FirstWorkStepHeadLine</td><td>FirstWorkStepDesc</td><td>documentTitle</td><td>eventDescriptionTitle</td><td>ePELTitle</td><td>ePELDescription</td><td></td></tr><tr><td>N777ZH</td><td>N777ZH_H-2202</td><td>WP.153455</td><td>3ML</td><td>GLF6</td><td>3ML - MINOR MAINTENANCE (FEB2022)</td><td>JCA.569398</td><td>3</td><td>8</td><td>JC</td><td>NULL</td><td>NULL</td><td>90001 - ARRIVAL / DEPARTURE</td><td>NULL</td><td>MHRS CAPTURING ON ARRIVAL / DEPARTURE</td><td>90001 - ARRIVAL / DEPARTURE</td><td>NULL</td><td>90001 - ARRIVAL / DEPARTURE</td><td>MHRS CAPTURING ON ARRIVAL / DEPARTURE</td><td></td></tr><tr><td>N777ZH</td><td>N777ZH_H-2202</td><td>WP.153455</td><td>3ML</td><td>GLF6</td><td>3ML - MINOR MAINTENANCE (FEB2022)</td><td>JCA.569399</td><td>3</td><td>9</td><td>JC</td><td>NULL</td><td>NULL</td><td>90002 - RAMP EQUIPMENT</td><td>NULL</td><td>MHRS CAPTURING ON RAMP EQUIPMENT</td><td>90002 - RAMP EQUIPMENT</td><td>NULL</td><td>90002 - RAMP EQUIPMENT</td><td>MHRS CAPTURING ON RAMP EQUIPMENT</td><td></td></tr><tr><td>N777ZH</td><td>N777ZH_H-2202</td><td>WP.153455</td><td>3ML</td><td>GLF6</td><td>3ML - MINOR MAINTENANCE (FEB2022)</td><td>JCA.569400</td><td>3</td><td>10</td><td>JC</td><td>NULL</td><td>NULL</td><td>90003 - AIDS FOR PRODUCTION (CONSUMABLE)</td><td>NULL</td><td>FOR PRODUCTION STAFF TO DEMANDCOMMERCIAL STORE ITEMS.  E.G. GLOVES/ SMALL RED BAGS ETC.</td><td>90003 - AIDS FOR PRODUCTION (CONSUMABLE)</td><td>NULL</td><td>90003 - AIDS FOR PRODUCTION (CONSUMABLE)</td><td>FOR PRODUCTION STAFF TO DEMANDCOMMERCIAL STORE ITEMS.  E.G. GLOVES/ SMALL RED BAGS ETC.</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555582</td><td>1</td><td>51</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-216000-05-1-02</td><td>NULL</td><td>DISCARD THE CABIN TEMPERATURE-SENSOR FILTERS</td><td>NULL</td><td>DISCARD THE CABIN TEMPERATURE-SENSOR FILTERS</td><td>9-216000-05-1-02</td><td>DISCARD THE CABIN TEMPERATURE-SENSOR FILTERS</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555589</td><td>1</td><td>58</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-236100-01-1-02</td><td>NULL</td><td>CHECK THE RESISTANCE FROM THE STATIC DISCHARGER TO THE BASE (RETAINER) AND FROM THE BASE (RETAINER) TO THE AIRCRAFT STRUCTURE - RH WING</td><td>NULL</td><td>CHECK THE RESISTANCE FROM THE STATIC DISCHARGER TO THE BASE (RETAINER) AND FROM THE BASE (RETAINER) TO THE AIRCRAFT STRUCTURE - RH WING</td><td>9-236100-01-1-02</td><td>CHECK THE RESISTANCE FROM THE STATIC DISCHARGER TO THE BASE (RETAINER) AND FROM THE BASE (RETAINER) TO THE AIRCRAFT STRUCTURE - RH WING</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555593</td><td>1</td><td>62</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-236100-02-1-03</td><td>NULL</td><td>GENERAL VISUAL INSPECTION OF STATIC DISCHARGER - HORIZONTAL AND VERTICAL STABILIZERS</td><td>NULL</td><td>GENERAL VISUAL INSPECTION OF STATIC DISCHARGER - HORIZONTAL AND VERTICAL STABILIZERS</td><td>9-236100-02-1-03</td><td>GENERAL VISUAL INSPECTION OF STATIC DISCHARGER - HORIZONTAL AND VERTICAL STABILIZERS</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555616</td><td>1</td><td>85</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-252331-01-2AC-01</td><td>NULL</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN L1-R1 AND L2-R2 DOORS) - CLEANING</td><td>NULL</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN L1-R1 AND L2-R2 DOORS) - CLEANING</td><td>9-252331-01-2AC-01</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN L1-R1 AND L2-R2 DOORS) - CLEANING</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555619</td><td>1</td><td>88</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-252331-01-2AC-04</td><td>NULL</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN CABIN ROW 39 AND END OF CABIN) - CLEANING</td><td>NULL</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN CABIN ROW 39 AND END OF CABIN) - CLEANING</td><td>9-252331-01-2AC-04</td><td>AREA BEHIND DADO PANELS (LOCATED BETWEEN CABIN ROW 39 AND END OF CABIN) - CLEANING</td><td></td></tr><tr><td>CGHKC</td><td>CGHKC_H-2201</td><td>WP.46536</td><td>AC</td><td>A333</td><td>AC - (943) C1+WIFI+CABIN MOD</td><td>CR.555622</td><td>1</td><td>91</td><td>CT</td><td>NULL</td><td>NULL</td><td>9-252344-02-1-04</td><td>NULL</td><td>CABIN "B" LH SIDE (FROM START OF ECONOMY TO L3 DOOR) - CHECK THAT ALL CABIN DECOMPRESSION PANELS ARE CLOSED</td><td>NULL</td><td>CABIN "B" LH SIDE (FROM START OF ECONOMY TO L3 DOOR) - CHECK THAT ALL CABIN DECOMPRESSION PANELS ARE CLOSED</td><td>9-252344-02-1-04</td><td>CABIN "B" LH SIDE (FROM START OF ECONOMY TO L3 DOOR) - CHECK THAT ALL CABIN DECOMPRESSION PANELS ARE CLOSED</td></tr></table>
      <button id="btn" onClick={() => fnExportToExcel('xlsx', 'myfirstsheet')}>
        Export to Excel
      </button>
    </div>
  );
}
export default App;

