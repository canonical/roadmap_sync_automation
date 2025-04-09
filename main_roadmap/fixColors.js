// provided by Maksim Beliaev
function replaceAllBackgroundColors() {
  var colorMap = {
    "#3d85c6": [  // Blue
      "#3c78d8", "#4a86e8", "#007aa6", "#0088cc", "#6fa8dc", "#4285f4", "#1155cc"
    ],
    "#cc0000": [  // Red
      "#990000", "#ff0000", "#a61c00", "#980000", "#e06666", "#ea4335"
    ],
    "#6aa84f": [  // Green
      "#38761d", "#429d35", "#93c47d", "#d9ead3"
    ],
    "#f3f3f3": [  // Grey
      "#b7b7b7", "#efefef", "#f8f8f8", "#f0f0f0"
    ],
    "#e59138": [  // Orange
      "#ff9900", "#e69138", "#f1c232", "#ffd966", "#fff2cc", "#f9cb9c", "#f6b26b"
    ]
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    if (!sheet.isSheetHidden()) { // Only process visible sheets
      var range = sheet.getDataRange();
      var backgrounds = range.getBackgrounds();
      
      for (var i = 0; i < backgrounds.length; i++) {
        for (var j = 0; j < backgrounds[i].length; j++) {
          let currentColor = backgrounds[i][j].toLowerCase();
          
          for (let targetColor in colorMap) {
            if (colorMap[targetColor].includes(currentColor)) {
              backgrounds[i][j] = targetColor;
              break; // Stop checking once we find a match
            }
          }
        }
      }
      
      range.setBackgrounds(backgrounds);
    }
  });
  
  SpreadsheetApp.flush();
}

