//set this to "off" when not doing sports ads
var sMode="off";






//imports all links from search result
function allImport() {

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
          var mode='nzcu';

        if(rng.getColumn()!=1)
        {Browser.msgBox("Put this link in column A adn retry"); return 0};
        
        
        var url=sheet.getActiveRange().getValue();
        
        
        
        var getRow=rng.getRow();
        var startRow=getRow;
        var row=getRow; 
        
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
        
        if(url.indexOf('overstock')>=0)
        {
           allImportOs(url)
           return 0
        
        }
        //for walmart
        

        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        var jsonData=getMyJsonSearch(html);
        var myItems=jsonData.preso.items
        var arr=[];
        var blankArr=["", "", "", "", "", "", "", ""];
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        for (var i in myItems)
        {
               var myItem=myItems[i];
               var itemNo=myItem.usItemId;
               var prodUrl="https://www.walmart.com"+myItem.productPageUrl;
               var wmTitle="";//myItem.title;
               wmTitle=replaceAll(wmTitle, "<mark>", "");
               wmTitle=replaceAll(wmTitle, "</mark>", "");
               var initial="";
               var date="";
               
               var sku="";
               var skugridVar="";
               
               var prevRow=row+2+arr.length-1;
               getRow=prevRow+1;
               var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
               var asin=repeatFrm;
        
               
               var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, prodUrl];
               arr.push(tempArr);
               
               
               
               var myVariant=myItem.variants;
               
               if(myVariant==undefined) // no variation
               {
                   arr.push(blankArr); arr.push(blankArr); arr.push(blankArr); //three more blank rows
               
                   continue;
                 }
               var myVariants=myVariant.variantData;
               
               for ( var i in myVariants)
               {
                 
                   arr.push(blankArr);
               
               }
               arr.push(blankArr); arr.push(blankArr); //two more blank rows
               
               
        
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);
   
}









function allImportOs(url)
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        var mode='nzcu';
       // var url="https://www.overstock.com/Pet-Supplies/Collars-Harnesses-Leashes/313/dept.html?sort=Avg.%20Customer%20Rating";
        
        
        
        var getRow=rng.getRow();
        var startRow=getRow;
        var row=getRow; 
        
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
 

        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        
        var n1=html.indexOf('window.__INITIAL_STATE__=')+('window.__INITIAL_STATE__=').length;
        var n2=html.indexOf('window.__HAS_RESULTS__=true;',n1);
        var n3=html.lastIndexOf("}",n2)+1;
        var html2=html.slice(n1,n3);
      //  GmailApp.sendEmail("sakib118.biz@gmail.com", "test", html2)
        var jsonData=JSON.parse(html2)
        var myItems=jsonData.products;
        var myItems=myItems[Object.keys(myItems)[0]].products;
        
        var arr=[];
        var blankArr=["", "", "", "", "", "", "", ""];
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        for (var i in myItems)
        {
               var myItem=myItems[i];
               var itemNo=myItem.sku;
               var prodUrl=myItem[Object.keys(myItem)[0]].productPage;
               var wmTitle="";//myItem.title;

               var initial="";
               var date="";
               
               var sku="";
               var skugridVar="";
               
               var prevRow=row+2+arr.length-1;
               getRow=prevRow+1;
               var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
               var asin=repeatFrm;
        
               
               var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, prodUrl];
               arr.push(tempArr);
               
               
               
               var myVariant=myItem.variants;
               
               if(myVariant==undefined) // no variation
               {
                   arr.push(blankArr); arr.push(blankArr); arr.push(blankArr); //three more blank rows
               
                   continue;
                 }
               var myVariants=myVariant.variantData;
               
               for ( var i in myVariants)
               {
                 
                   arr.push(blankArr);
               
               }
               arr.push(blankArr); arr.push(blankArr); //two more blank rows
               
               
        
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);


}
















function importOneByOne()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var rng=sheet.getActiveRange();
    var sr=rng.getRow();
    var er=rng.getLastRow();
       
    for (var i=sr; i<=er; i++)
    {
          if(sheet.getRange(i, 8).getValue()==""){continue};
          
          if(sheet.getRange(i, 1).getValue()!=""){continue;} //skip if already imported
          sheet.getRange(i, 8).activate();

          var n= importFromSource1("m");
          imShowSideBar()
         // i=i+n; //increase by number of variations
    
    
    }

}



function makeBatchAds()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var rng=sheet.getActiveRange();
    var sr=rng.getRow();
    var er=rng.getLastRow();
       
    for (var i=sr; i<=er; i++)
    {
          if(sheet.getRange(i, 8).getValue()==""){continue};
          sheet.getRange(i, 8).activate();
          try
          {
              nfl_womensTee();
          }
          
          catch(err)
          {
              continue;
          }
          //nfl_elf();
          
          
    
    }

    


}










function batchAdThisRow() {
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













function findTeam(title, col1, col2)
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet1=ss.getSheetByName('Sports Bedding');
    var values1=sheet1.getDataRange().getValues();
    
    
    var sheet2=ss.getSheetByName('Mapping3');
    var values2=sheet2.getDataRange().getValues();
    
    for (var i=1; i<values1.length; i++)
    {
        var tempTeamName=values1[i][col1-1];
        if(title.indexOf(tempTeamName)>=0)
        {
              break;
        
        }
        
    
    }
    var team=values1[i][col1-1]; 
    var color=values1[i][col1-1+1]; //column to right
    var longName=values1[i][col1-3]; 
    
    var flag=0;
    var material="";
    for (var i=1; i<values2.length; i++)
    {
        var tempMaterial=values2[i][1-1];
        if(title.indexOf(tempMaterial)>=0)
        {
              flag=1; break;
        
        }
        
    
    }
    
    
    if(flag==1)
     { material=values2[i][1-1];} 
    
     return [team, color, longName, material]
      


}






function sportsAdd() 
{
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      
      var color=details[1];
      
      var material=details[3];
      
      var partsTitle=sourceTitle.split(" x ");
      
      var size1=59;//partsTitle[0].match(/\d+/)[0];
      var size2=6.5;//partsTitle[1].match(/\d+/)[0];
      
      var amTitle='1 piece Nfl '+name+' Adult Big Logo Scarf 59 x 6.5 Inches, Football Themed Fashion Accessory Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan '+color+', Acrylic'
      var b1='1 piece Nfl '+fullName+' Adult Big Logo Scarf 59 x 6.5 Inches, Football Themed Fashion Accessory Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan '+color+', Acrylic'
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="1 Nfl Scarf";
      var b4=' This 100 percent Acrylic, cozy full length scarf is officially licensed by the NFL';
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 18).setValue(b4);
      
var img=imShowSideBar(); //showWmImages(); 
      var imFrm='=IMAGE("'+img+'", 1)';
      sheet.getRange(row, 5).setValue(imFrm);
     /// sheet.getRange(row, 18).setValue('This throw can be used out at a game, on a picnic, in the bedroom, or cuddled under in the den while watching the game on TV.')
      sheet.getRange(row, 20).setValue(img);
      
      sheet.getRange(row, 2).setValue("NZCU4AS");
      var today= Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
      sheet.getRange(row, 3).setValue(today);
      
      Logger.log(amTitle)
      var a=10

  
}












