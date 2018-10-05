

function nflFootBall() {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1=19;//partsTitle[0].match(/\d+/)[0];
      var size2=30; //partsTitle[1].match(/\d+/)[0];
      
      var amTitle='9 Inch NFL '+name+ ' Football Composite Leather, Brown Color Black Laser Stamped Team Logo Sports Themed, Gift For Fan Collectible Athletic Spirit' ;
      
      var b1='9 Inch NFL '+fullName+ ' Football Composite Leather, Brown Color Black Laser Stamped Team Logo Sports Themed, Gift For Fan Collectible Athletic Spirit' ;
      var includes="Distressed brown color and black stitching helps football stand out."
      var dim="Football: 9 inches long"
      var b4="Composite leather stitching and laser stamped team logo.";
      
  //    sheet.getRange(row, 12).setValue(amTitle);
 ////     sheet.getRange(row, 34).setValue(b1);
 //     sheet.getRange(row, 16).setValue(includes);
    //  sheet.getRange(row, 17).setValue(dim);
      var img=imShowSideBar()
     //  sheet.getRange(row, 18).setValue(b4);
    //   sheet.getRange(row, 20).setValue(img);
   //  sheet.getRange(row, 2).setValue('NZCU4NTC');
   //   sheet.getRange(row, 3).setValue(new Date());
     

  
}














function nflTableCloth() {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1=19;//partsTitle[0].match(/\d+/)[0];
      var size2=30; //partsTitle[1].match(/\d+/)[0];
      
      var amTitle='54 x 102 Inch NFL '+name+' Tablecloth, Football Themed Rectangle Table Cover Sports Patterned, Team Color Logo Fan Merchandise Athletic Spirit '+color+', Plastic ';
      var b1='54 x 102 Inch NFL '+fullName+' Tablecloth, Football Themed Rectangle Table Cover Sports Patterned, Team Color Logo Fan Merchandise Athletic Spirit '+color+', Plastic ';
      var includes="Get your game day festivities started by covering your party table with the NFL table cover."
      var dim="Tablecloth Dimension: 54 Inches x 102 Inches"
      var b4="Beautiful blue tablecloth features a nfl pattern and design.";
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      var img=showWmImages();
       sheet.getRange(row, 18).setValue(b4);
       sheet.getRange(row, 20).setValue(img);
      sheet.getRange(row, 2).setValue('NZCU4NTC');
      sheet.getRange(row, 3).setValue(new Date());
     

  
}


















function nflMats() {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1=19;//partsTitle[0].match(/\d+/)[0];
      var size2=30; //partsTitle[1].match(/\d+/)[0];
      
      var amTitle=size1+'" x '+size2+'" Inch NFL '+name+' Door Mat Set Printed Logo, Football Themed Sports Patterned Bathoroom Kitchen Outdoor Carpet Rug Gift for Fan Merchandise Athletic Team Spirit '+color +', Nylon';
      var b1='19" x 30" Inch NFL '+fullName+' Door Mat Printed Logo, Football Themed Sports Patterned Bathoroom Kitchen Outdoor Carpet Area Rug Gift for Fan Merchandise Athletic Team Spirit '+color+', Nylon';
      var includes="Includes: 1 NFL Carpet Mat"
      var dim="Dimensions: 17 x 26 inches"
      var b4="Non-skid recycled vinyl backing";
      
      //sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      //sheet.getRange(row, 16).setValue(includes);
      //sheet.getRange(row, 17).setValue(dim);
      //var img=showWmImages();
     // sheet.getRange(row, 18).setValue(b4)
     // sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}



function nflFootballMats() {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1=19;//partsTitle[0].match(/\d+/)[0];
      var size2=30; //partsTitle[1].match(/\d+/)[0];
      
      var amTitle='22" x 35" NFL '+name+' Floor Mat Printed Logo, Football Shaped Area Rug Oval Rug Sports Patterned Themed Gift for Fan Merchandise Athletic Team Spirit, Brown '+color+', Nylon';
      var b1='22" x 35" NFL '+fullName+' Floor Mat Printed Logo, Football Shamped Area Rug Oval Rug Sports Patterned Themed Gift for Fan Merchandise Athletic Team Spirit, Brown '+color+',, Nylon';
      //var includes="Includes: 1 NFL football Mat"
      //var dim="Dimensions: 22 x 35 inches"
      //var b4="Non-skid recycled vinyl backing";
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
   //   sheet.getRange(row, 16).setValue(includes);
    //  sheet.getRange(row, 17).setValue(dim);
      //var img=imShowSideBar()
      //sheet.getRange(row, 18).setValue(b4)
    //  sheet.getRange(row, 20).setValue(img)
      


  
}