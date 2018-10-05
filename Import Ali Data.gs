
function determineIphoneClass()
{

            var rng =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
            var editedSheet = rng.getSheet();
            var getRow = rng.getRow();
            var getCol = rng.getColumn();
            var getURL = rng.getValue();
            var sheet=editedSheet
            var arrVals=rng.getValues();
            var ss=SpreadsheetApp.getActiveSpreadsheet();
            var arr=[]
            var colVs=[];
            var B2s=[];
            var B4s=[];
            var B5s=[];
//-----------------------------------------------------------------

            if(arrVals[0][0].toString().toLowerCase().indexOf("iphone")<0){return 0}
            var iphoneSizes=ss.getSheetByName("Dropdown Options").getRange("H1:P").getValues();
            
            
            var myColor="#9fc5e8"; //light blue 2
            //white="#ffffff"
            var colors=[[myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor]];
            for (var i=0; i<arrVals.length; i++)
            {
               var vari=arrVals[i][6];
               arr.push(["",""]);
               colVs.push([""]);
               B2s.push([""]);
               B4s.push([""]);
               B5s.push([""]);
               if(vari.toString().indexOf("|")>0)
               {
                   vari=vari.split("|")[1];
                   var thisColor=arrVals[i][6].split("|")[0];
                   if(i>0)
                   {
                       var prevColor=arrVals[i-1][6].split("|")[0];
                       Logger.log(thisColor+"  "+prevColor)
                       if(thisColor!=prevColor)
                       {
                                              Logger.log("inside    "+thisColor+"  "+prevColor)

                           if(myColor=="#9fc5e8"){myColor="#ffffff"}
                           else{
                             myColor="#9fc5e8";
                           }
                       }
                       colors.push([myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor]);

                   }

                   
                   for(var j=0; j<iphoneSizes.length; j++)
                   {
                           var varTemp=iphoneSizes[j][0].toString().toLowerCase();
                         if(vari.toString().toLowerCase().indexOf(varTemp)>=0 || vari.indexOf(iphoneSizes[j][0])>=0)
                         {
                                 Logger.log(i+"  "+j)
                                 arr[i][0]=(iphoneSizes[j][2]).replace("PLUS","PLUS_Sized_Case_Bigger_Screen");
                                 arr[i][1]=iphoneSizes[j][3];
                                 colVs[i][0]=iphoneSizes[j][5];
                                 B2s[i][0]=iphoneSizes[j][6];
                                 B4s[i][0]=iphoneSizes[j][7];
                                 B5s[i][0]=iphoneSizes[j][2]+"__";
                                 
                                 break
                                 
                         }
                   
                   }
                 
               }
            
            }
            
            
            
            
            
            sheet.getRange(getRow, 36, arr.length, arr[0].length).setValues(arr);
            sheet.getRange(getRow, 1, arr.length,colors[0].length).setBackgrounds(colors);
            
            sheet.getRange(getRow, 22, arr.length, 1).setValues(colVs);
            sheet.getRange(getRow, 16, arr.length, 1).setValues(B2s); 
            sheet.getRange(getRow, 18, arr.length, 1).setValues(B4s); 
            sheet.getRange(getRow, 38, arr.length, 1).setValues(colVs);   
            //sheet.getRange(getRow, 39, arr.length, 1).setValues(B5s);   
            
            
            
            var a=10
  }














function  myitem(url)
{
            
            var getURL=url;
            getURL=getURL.slice(0, getURL.indexOf('.html')+5);
            
            var n1=getURL.indexOf('.html');
            var n2=getURL.lastIndexOf('/');
            
            var itemNo=getURL.slice(n2+1,n1);

            return itemNo;


}


function alexpressImportChecker()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet(); 
    var sheet=ss.getSheetByName("LINKS**");
    var values1=sheet.getRange("A100:A").getValues();
    var values2=sheet.getRange("B100:B").clearContent().getValues();
    var values3 = sheet.getRange("C100:C").clearContent().getValues();
              var sheets=ss.getSheets();
              var allItemNos=[];
              var sheetNames=[];
              var statuses=[];
          for (var s=0; s<sheets.length; s++)
          {
              
              if(isLister(sheets[s])==false) {continue}
              var itemnos=sheets[s].getRange("F1:F").getValues().join("|").split("|")
              var status=sheets[s].getRange("K1:K").getValues().join("|").split("|");
              allItemNos.push(itemnos);
              statuses.push(status);
              sheetNames.push(sheets[s].getName())
           }
             
    for (var i=0; i<values1.length; i++)
    {
          var itemNo=values1[i][0];
          if (itemNo==""){continue}
          
          for (var j=0; j<allItemNos.length; j++)
          {
              if(allItemNos[j].indexOf(itemNo.toString())>=0)
              {
                  values2[i][0]=sheetNames[j];
                  values3[i][0] = statuses[j][allItemNos[j].indexOf(itemNo.toString())]
              }
              
              else if (allItemNos[j].indexOf(itemNo)>=0)
              {
                  values2[i][0]=sheetNames[j];
                  values3[i][0] = statuses[j][allItemNos[j].indexOf(itemNo)] 
              
              }
              
          
          }
          
    
    
    }
    
    sheet.getRange("B100:B").setValues(values2);
    sheet.getRange("C100:C").setValues(values3);
    

}










function makeCSV4UploadAE()
{

      var ss=SpreadsheetApp.getActiveSpreadsheet()
      var csvSheet=ss.getSheetByName('CSV');
      csvSheet.getRange(2, 1, csvSheet.getMaxRows()-1, csvSheet.getMaxColumns()).clearContent();
      var sheets=ss.getSheets();
      var arrs=[];
     
   for(var s=0; s<sheets.length; s++)      
   {   
                
                
                
                var sheet=sheets[s];
                var lr=sheet.getLastRow();
                lr=lr<6?6:lr;
                
                if(isLister(sheet)==false){continue};
                var values=sheet.getRange(6,1,lr-6+1, 52).getValues();
                
                for (var i= 10; i<values.length; i++)
                {
                      var status=values[i][11-1];
                      var init=values[i][2-1];
                      var prevInit=values[i][4-1]
                      
                      if(status=="COMPLETE")
                      {
                           if(init.indexOf("AEG2")>0)
                           {
                                  values[i][4]=sheet.getName()+"|"+(i+1);
                                  for( var j=11; j<values[i].length, isNaN(values[i][j]) ; j++)
                                  {
                                    values[i][j]=values[i][j].replace("=","'=")
                                  }
                                  arrs.push(values[i]);
                            }
                            
                      
                      }
                
                }//end of i for 

   }


    csvSheet.getRange(2, 1, arrs.length, 52).setValues(arrs)




}












function importAliData(rng,spc)
{
  var rng =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  var editedSheet = rng.getSheet();
  var getRow = rng.getRow();
  var getCol = rng.getColumn();
  var getURL = rng.getValue();
   

     var mode="live";
            
            
  
  
  
  if (getCol == 8 && isLister(editedSheet)) 
  {
            
            
            if(getURL.indexOf("aliexpress")<0){return 0;}
            getURL=getURL.slice(0, getURL.indexOf('.html')+5);
            
            var n1=getURL.indexOf('.html');
            var n2=getURL.lastIndexOf('/');
            
            var itemNo=getURL.slice(n2+1,n1).toString();;
            if(itemNo.indexOf("_")>0){
              itemNo=itemNo.split("_")[1];
            }
            
            if(spc==""){
                        var ssDoneAli=SpreadsheetApp.openById("1UN-nTLNpO7Y9vooVJMIh1YPDzfI1p8NdG3Ryje08R6E");
                        var values1=ssDoneAli.getSheetByName("DB").getRange("A1:A").getValues().join(",")+"'"+ssDoneAli.getSheetByName("Item Checker").getRange("A1:A").getValues().join(",");
                        if(values1.indexOf(itemNo.toString())>=0){
                            Browser.msgBox("This item is already done by someone else. Delete It or Consult Nazmus"); return 0
                            
                        } else {
                          var sheetDoneAli= ssDoneAli.getSheetByName("Item Checker");
                          var lrAliDone=last_row(sheetDoneAli,1);
                          
                          var existingAliValues=sheetDoneAli.getRange(1,1,sheetDoneAli.getMaxRows(), sheetDoneAli.getMaxColumns()).getValues();
                          
                          for(var si=10; si<existingAliValues.length; si++){
                            if(existingAliValues[si][0]==""){break}
                          
                          }
                          
                          sheetDoneAli.getRange(si+1, 1).setValue(itemNo);
                          sheetDoneAli.getRange(si+1, 4).setValue("Nazmus");
                          sheetDoneAli.getRange(si+1, 5).setValue(Utilities.formatDate(new Date(), "UTC", "YYYY-MM-dd"));
                          var lrr=(si+1);
                          Browser.msgBox("Item no added to row "+ lrr +" on Aliexpress Products sheet");
                        }
             }   
            
            
            
            
            
            
            
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var sheet = ss.getActiveSheet();
            var option = {
              'muteHttpExceptions' : true
            };
            var html = UrlFetchApp.fetch(getURL, option).getContentText();
            var htmlOrig=html;
            
            //var sheetUPC=ss.getSheetByName("UPC");
            //var lrUPC=sheetUPC.getLastRow();
            
            var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
            
            var arrVals = [];
            var arrVals2=[];
            var arrVals3=[]
            var m1 = html.indexOf('<h1 class="product-name" itemprop="name">')+('<h1 class="product-name" itemprop="name">').length;
            var m2 = html.indexOf("</h1>", m1);
            var title = html.slice(m1, m2);
            
            
            var p1=html.indexOf("var skuProducts=[")+('var skuProducts=[').length;
            var p2=html.indexOf("];",p1);
            var arrProd=html.slice(p1,p2).split('}},');
            var currentValues=sheet.getRange(getRow, 1, arrProd.length,1).getValues()
            
            
            
            
            
            
            
            var initials="";
            var today=Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
            var asin="";
            var sellerSku="";
            var  upc="";
             var comment=allLenFrm;
            var status="";
            var prodName="";
            var type="";
            var type='BedAndBath';
            var catagory="";
            
            
            var b2="";
            var b3="";
            var b4="**SHIPS From China - 10-14 Business Days**";
            var terms="";
            var imageMain="";
            var color="";
            var size="";
            var material="";
            var b1="";
            
            
            
            
            
            var skuProps=html.split('data-sku-prop-id'); //split by itemprop
            var imgTitles=[];
            var imgIds=[];
            var countryTitles=[];
            var countryIds=[];
            var sizeTitles=[];
            var sizeIds=[];
            var materialTitles=[];
            var materialIds=[];
            var specTitles=[];
            var specIds=[];
            var skippedRows=0;
            
            
            for (var i=1; i<skuProps.length; i++)
            {
                              if(skuProps[i].indexOf('="14"')==0)  //image item pro is found
                              {
                                      var n1=html.indexOf('<li class="item-sku-image">');
                                      
                                      var n2=html.indexOf('</ul>',n1);
                                      var imageText=html.slice(n1,n2).trim();
                                      var imageTextArr=imageText.split('<li class="item-sku-image">');
                                      for (var j=1; j<imageTextArr.length; j++)
                                      {
                                          var imageLarge=imageTextArr[j];
                                          var n1=imageLarge.indexOf('title="')+7;
                                          var n2=imageLarge.indexOf(' href="javascript',n1);
                                          var imTitle=imageLarge.slice(n1,n2-1);
                                          imgTitles.push(imTitle);
                                          
                                          var n1=imageLarge.indexOf('data-sku-id="')+13;
                                          var n2=imageLarge.indexOf('id',n1);
                                          
                                          var imgId=imageLarge.slice(n1,n2-2);
                                          imgIds.push(imgId);
                                          
                                      }
                              
                              
                              
                              }//end of image array prod id=14
                              
                              else if(skuProps[i].indexOf('="200007763')==0)  //country item prop is found
                              {
                                      var n1=skuProps[i].indexOf('<li>');
                                      var n2=skuProps[i].indexOf('</ul>',n1);
                                      var countryText=skuProps[i].slice(n1,n2).trim();
                                      var countryTextArr=countryText.split('<li>');
                                      for (var j=1; j<countryTextArr.length; j++)
                                      {
                                          var countryLarge=countryTextArr[j];
                                          countryLarge=replaceAll(countryLarge,"<span>","");
                                          countryLarge=replaceAll(countryLarge,"</span>","")
                                          var n1=countryLarge.indexOf('data-code="')+11;
                                          var n2=countryLarge.indexOf('>',n1);
                                          var n3=countryLarge.indexOf('<',n2);
                                          var countryTitle=countryLarge.slice(n2+1,n3);
                                          countryTitles.push(countryTitle);
                                          
                                          var n1=countryLarge.indexOf('data-sku-id="')+13;
                                          var n2=countryLarge.indexOf('id',n1);
                                          
                                          var countryId=countryLarge.slice(n1,n2-2);
                                          countryIds.push(countryId);
                                          
                                      }
                              
                              
                              
                              }//end of image array prod id=country
                              
                              else if(skuProps[i].indexOf('="5"')==0 )  //country item prop is found
                              {
                                      var n1=skuProps[i].indexOf('<li>');
                                      var n2=skuProps[i].indexOf('</ul>',n1);
                                      var sizeText=skuProps[i].slice(n1,n2).trim();
                                      var sizeTextArr=sizeText.split('<li>');
                                      for (var j=1; j<sizeTextArr.length; j++)
                                      {
                                          var sizeLarge=sizeTextArr[j];
                                          sizeLarge=replaceAll(sizeLarge,"<span>","");
                                          sizeLarge=replaceAll(sizeLarge,"</span>","")
                                          var n1=sizeLarge.indexOf('data-code="')+11;
                                          var n2=sizeLarge.indexOf('>',n1);
                                          var n3=sizeLarge.indexOf('<',n2);
                                          var sizeTitle=sizeLarge.slice(n2+1,n3);
                                          sizeTitles.push(sizeTitle);
                                          
                                          var n1=sizeLarge.indexOf('data-sku-id="')+13;
                                          var n2=sizeLarge.indexOf('id',n1);
                                          
                                          var sizeId=sizeLarge.slice(n1,n2-2);
                                          sizeIds.push(sizeId);
                                          
                                      }
                              
                              
                              
                              }//end of image array prod id=size
                              
                               else if(skuProps[i].indexOf('="10"')==0 )  //MATERIAL item prop is found
                              {
                                      var n1=skuProps[i].indexOf('<li>');
                                      var n2=skuProps[i].indexOf('</ul>',n1);
                                      var materialText=skuProps[i].slice(n1,n2).trim();
                                      var materialTextArr=materialText.split('<li>');
                                      for (var j=1; j<materialTextArr.length; j++)
                                      {
                                          var materialLarge=materialTextArr[j];
                                          materialLarge=replaceAll(materialLarge,"<span>","");
                                          materialLarge=replaceAll(materialLarge,"</span>","")
                                          var n1=materialLarge.indexOf('data-sku-id')+11;
                                          var n2=materialLarge.indexOf('>',n1);
                                          var n3=materialLarge.indexOf('<',n2);
                                          var materialTitle=materialLarge.slice(n2+1,n3);
                                          materialTitles.push(materialTitle);
                                          
                                          var n1=materialLarge.indexOf('data-sku-id="')+13;
                                          var n2=materialLarge.indexOf('id',n1);
                                          
                                          var materialId=materialLarge.slice(n1,n2-2);
                                          materialIds.push(materialId);
                                          
                                      }
                              
                              
                              }//end of image array prod id=size
                              
                              
                               
                              else if(skuProps[i].indexOf('="183"')==0 )  //specification prop id
                              {
                                      var n1=skuProps[i].indexOf('<li>');
                                      var n2=skuProps[i].indexOf('</ul>',n1);
                                      var specText=skuProps[i].slice(n1,n2).trim();
                                      var specTextArr=specText.split('<li>');
                                      for (var j=1; j<specTextArr.length; j++)
                                      {
                                          var specLarge=specTextArr[j];
                                          specLarge=replaceAll(specLarge,"<span>","");
                                          specLarge=replaceAll(specLarge,"</span>","")
                                          var n1=specLarge.indexOf('data-sku-id')+11;
                                          var n2=specLarge.indexOf('>',n1);
                                          var n3=specLarge.indexOf('<',n2);
                                          var specTitle=specLarge.slice(n2+1,n3);
                                          specTitles.push(specTitle);
                                          
                                          var n1=specLarge.indexOf('data-sku-id="')+13;
                                          var n2=specLarge.indexOf('id',n1);
                                          
                                          var specId=specLarge.slice(n1,n2-2);
                                          specIds.push(specId);
                                          
                                      }
                              
                              
                              }
                              
                              
     
            }
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            var vFrm=''
            var vFrmFlag=0;
            var sheetLister=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lister")
            var profit=sheetLister.getRange("U5").getFormulaR1C1();
            //var amPrice=sheetLister.getRange("R5").getFormulaR1C1();
            var amPrice="=ROUNDUP((R[0]C[17]-(-(R[0]C[16])+(R[0]C[16]*0)))/0.85)-0.01"; //"=ROUND((R[0]C[17]-(-(R[0]C[16])+((R[0]C[16]*0.12))+((R[0]C[16])-(R[0]C[16]*0.12))*0.0688))/0.85,0)-0.01";
            var skippedRows=0;
            var arr=[]; //for iphone
            var colVs=[];
            var B2s=[];
            var B4s=[];
            
            
            for (var i=0; i<arrProd.length; i++)
            {
            
                
                  
                
                
                
                 var skipFlag=0;

                 var text=arrProd[i]+"}}"
                 text=text.replace("}}}}",'}}'); //for last element
                 var jsonData=JSON.parse(text);
                 var variationLong=jsonData.skuAttr;
                 if(variationLong==undefined){var variation="";}
                 else
                 {
                             var skuAttrAll=variationLong.split(";");
                             var partVars=[];
                             for (var a=0; a<skuAttrAll.length; a++)
                             {
                                   var skuAttr=skuAttrAll[a];
                                   if(skuAttr.indexOf('#')>=0)
                                   {
                                     skuAttr=skuAttr.slice(0,skuAttr.indexOf('#'));
                                     }
                                   var property=skuAttr.split(":")[0];
                                   var value=skuAttr.split(":")[1];
                                   
                                   if(property=="14" && imgTitles.length>0)
                                   {
                                           var thisTitle=imgTitles[imgIds.indexOf(value)];
                                           partVars.push(thisTitle);
                                   }
                                   
                                   
                                   
                                   if(property=="200007763" && countryTitles.length>0)
                                   {
                                           var thisCountry=countryTitles[countryIds.indexOf(value)];
                                           if (thisCountry!="China")
                                           {
                                              skipFlag=1;
                                              break
                                           }
                                           partVars.push(thisCountry);
                                   }
                                   
                                   
                                   
                                    
                                   if(property=="5" && sizeTitles.length>0)
                                   {
                                           var thisSize=sizeTitles[sizeIds.indexOf(value)];
                                           partVars.push(thisSize);
                                   }
                                   
                                     if(property=="10" && materialTitles.length>0)
                                   {
                                           var thisMaterial=materialTitles[materialIds.indexOf(value)];
                                           partVars.push(thisMaterial);
                                   }
                                   
                                     if(property=="183" && specTitles.length>0)
                                   {
                                           var thisSpec=specTitles[specIds.indexOf(value)];
                                           partVars.push(thisSpec);
                                   }
                             
                             
                             }// for a loop
                             
                             var variation=partVars.join("|");
                 }
                 
                 
                 
                 var price=jsonData.skuVal.actSkuCalPrice
                 if(price==undefined)
                 {
                   price=jsonData.skuVal.skuCalPrice;
                 }
                 
                  var qty=jsonData.skuVal.inventory;
                   //if(skipFlag==1){continue}
                   Logger.log(qty+"  "+skipFlag)
                   if(qty<10 || skipFlag==1)
                   {
                         skippedRows++;
                         continue;
                         }
                 
                 var rowTemp=i+getRow-skippedRows;
                 var priceWithShipping='='+price+"+R[0]C[2]";
                      var lenFrm="=LEN(L"+rowTemp+")"
                      var lenFrm2="=LEN(S"+rowTemp+")"
                      var lenFrm3="=LEN(R"+rowTemp+")"
                      var imFrm2="=R[0]C[-4]";
                      var imFrm3="=R[0]C[-5]";
                      var imFrm4="=R[0]C[-6]";
                 
                 var mTerms="0";
                 
                 initials="=R[-1]C[0]"
                 mTerms="=R[-1]C[0]"
                 
                 
                 if(vFrmFlag==1)
                 {
                              prodName=vFrm;
                              type=vFrm;
                              catagory=vFrm;
                              b2=vFrm;
                              b3=vFrm;
                              b4=vFrm;
                              terms=vFrm;
                              material=vFrm;
                              size=vFrm;
                              initials="=R[-1]C[0]"
                              mTerms="=R[-1]C[0]"
                 
                 }
                 
                 
                 
                 arr.push(["",""]);
                 colVs.push([""]);
                 B2s.push([""]);
                 B4s.push([""]);
                 
                 vFrmFlag=1;
               //  title=title.replace('PLUS', "PLUS Sized Case Bigger Screen");
                 arrVals.push([title, initials, today, asin, sellerSku, itemNo, variation, getURL, upc, comment,status, prodName,type, catagory,amPrice, b2, b3, b4, terms, imageMain, color,size,material, imFrm2, imFrm3, imFrm4,lenFrm3,lenFrm2,lenFrm,qty, priceWithShipping, profit, mTerms,b1]);
                 
                 
            
            }
            
            function sortFunction(a, b) {
                if (a[6] === b[6]) {
                    return 0;
                }
                else {
                    return (a[6] < b[6]) ? -1 : 1;
                }
            }
            
            
            arrVals =arrVals.sort(sortFunction)            
            
            
            
            
            
                        
            var currentValues=sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length).getValues()
            for (var r=1; r<currentValues.length; r++)
            {
                Logger.log("row "+ r)
                for (var l=0; l<currentValues[r].length; l++)
                {
                  Logger.log("row "+ l)
                  if(currentValues[r][l]!="")
                  {
                       Browser.msgBox("This will overwrite value in "+sheet.getRange(getRow+r, l+1).getA1Notation());
                       return 0
                  }
                }
            
            }
            

            
            
            
         //   arrVals =arrVals.sort(sortFunction)            
            
            sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length).setValues(arrVals);
           // deactivateFormulas(sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length));
            //-----------------------------------------------------------------
            sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length).activate();
            
            if(arrVals[0][0].toString().toLowerCase().indexOf("iphone")<0){return 0}
            
            determineIphoneClass();
            setAliImages();
            return 0;
            
            var iphoneSizes=ss.getSheetByName("Dropdown Options").getRange("H1:O").getValues();
            
            var myColor="#9fc5e8"; //light blue 2
            //white="#ffffff"
            var colors=[[myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor]];
            for (var i=0; i<arrVals.length; i++)
            {
               var vari=arrVals[i][6];
               
               if(vari.toString().indexOf("|")>0)
               {
                   vari=vari.split("|")[1];
                   var thisColor=arrVals[i][6].split("|")[0];
                   if(i>0)
                   {
                       var prevColor=arrVals[i-1][6].split("|")[0];
                       Logger.log(thisColor+"  "+prevColor)
                       if(thisColor!=prevColor)
                       {
                                              Logger.log("inside    "+thisColor+"  "+prevColor)

                           if(myColor=="#9fc5e8"){myColor="#ffffff"}
                           else{
                             myColor="#9fc5e8";
                           }
                       }
                       colors.push([myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor,myColor]);

                   }

                   
                   for(var j=0; j<iphoneSizes.length; j++)
                   {
                           var varTemp=iphoneSizes[j][0].toString().toLowerCase();
                         if(vari.toString().toLowerCase().indexOf(varTemp)>=0 || vari.indexOf(iphoneSizes[j][0])>=0)
                         {
                                 arr[i][0]=iphoneSizes[j][2];
                                 arr[i][1]=iphoneSizes[j][3];
                                 colVs[i][0]=iphoneSizes[j][5];
                                 B2s[i][0]=iphoneSizes[j][6];
                                 B4s[i][0]=iphoneSizes[j][7];
                                 break
                                 
                         }
                   
                   }
                 
               }
            
            }
            
            
            
            
            
            sheet.getRange(getRow, 36, arr.length, arr[0].length).setValues(arr);
            sheet.getRange(getRow, 1, arr.length,colors[0].length).setBackgrounds(colors);
            
            sheet.getRange(getRow, 22, arr.length, 1).setValues(colVs);
            sheet.getRange(getRow, 16, arr.length, 1).setValues(B2s); 
            sheet.getRange(getRow, 18, arr.length, 1).setValues(B4s); 
            sheet.getRange(getRow, 38, arr.length, 1).setValues(colVs);   

            
            
            
            var a=10
       }
}













function showAliImages()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
       
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                    var n1=html.indexOf('data-sku-prop-id="14"');
                    
                    if(n1==-1)
                    {
                          showAliImagesNoVar(html);
                          return 0;
                    }
                    
                    var n2=html.indexOf('<li class',n1);
                    var n3=html.indexOf('</ul>',n2);
                    var imageText=html.slice(n2,n3).trim();
                    var imageTextArr=imageText.split("</li>")
                 



                    var sbHtml='<br>';
                    var htmlArr=[];
                    
                    var images=imageTextArr;
                    for (var i=0; i<images.length; i++) 
                    {
                             var imageLarge=images[i];
                             var n1=imageLarge.indexOf('title="')+7;
                             var n2=imageLarge.indexOf('"',n1);
                             var color=imageLarge.slice(n1,n2);
                             var n1=imageLarge.indexOf('bigpic="')+8;
                             var n2=imageLarge.indexOf('"/',n1);
                             var imUrl=imageLarge.slice(n1,n2)
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/"+imUrl;
                             var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/"+imUrl;
                             var imageURL=imUrl;//"http://res.cloudinary.com/demo/image/fetch/"+imUrl;



                            var tempHtml='<br>Color: '+color+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'

                          +'<form>'
                           //+'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                           // +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                            +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'

                           +'</form>'
                           +'<br><hr>'
                          
                          htmlArr.push(tempHtml);
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                    
                    }
        
        
        
        
                 //  htmlArr=htmlArr.sort();
        
                   sbHtml=sbHtml+htmlArr.join('<br>');
                                
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }




}





function showAliImagesNoVar(html)

{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
       
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    //var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                    var n1=html.indexOf('<ul class="image-thumb-list" id="j-image-thumb-list">');
                    
                   
                    
                    var n2=html.indexOf('<li>',n1);
                    var n3=html.indexOf('</ul>',n2);
                    var imageText=html.slice(n2,n3).trim();
                    var imageTextArr=imageText.split("</li>")
                 



                    var sbHtml='<br>';
                    var htmlArr=[];
                    
                    var images=imageTextArr;
                    for (var i=0; i<images.length-1; i++) 
                    {
                             var imageLarge=images[i];
                             
                             var color=""
                             var n1=imageLarge.indexOf('src="')+5;
                             var n2=imageLarge.indexOf('.jpg',n1)+4;
                             var imUrl=imageLarge.slice(n1,n2)
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/"+imUrl;
                             var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/"+imUrl;
                             var imageURL=imUrl;//"http://res.cloudinary.com/demo/image/fetch/"+imUrl;



                            var tempHtml='<br>Color: '+color+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'

                          +'<form>'
                           // +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                           // +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                            +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'

                           +'</form>'
                           +'<br><hr>'
                          
                          htmlArr.push(tempHtml);
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                    
                    }
        
        
        
        
                 //  htmlArr=htmlArr.sort();
        
                   sbHtml=sbHtml+htmlArr.join('<br>');
                                
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }




}



function setAliImages()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
       
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValues()[0][7].toString(); //column H
                    if(getURL==""){return 0};
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                    var n1=html.indexOf('data-sku-prop-id="14"');
                    
                    if(n1==-1)
                    {
                          showAliImagesNoVar(html);
                          return 0;
                    }
                    
                    var n2=html.indexOf('<li class',n1);
                    var n3=html.indexOf('</ul>',n2);
                    var imageText=html.slice(n2,n3).trim();
                    var imageTextArr=imageText.split("</li>")
                 



                    var sbHtml='<br>';
                    var htmlArr=[];
                    
                    var images=imageTextArr;
                    var allColors=[];
                    var allUrls=[];
                    var values=sheet.getActiveRange().getValues();

                    var allImages=sheet.getRange(sheet.getActiveRange().getRow(), 20, values.length, 1).getValues();
                    for (var i=0; i<images.length; i++) 
                    {
                             var imageLarge=images[i];
                             var n1=imageLarge.indexOf('title="')+7;
                             var n2=imageLarge.indexOf('"',n1);
                             var color=imageLarge.slice(n1,n2);
                             var n1=imageLarge.indexOf('bigpic="')+8;
                             var n2=imageLarge.indexOf('"/',n1);
                             var imUrl=imageLarge.slice(n1,n2)
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/"+imUrl;
                             var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/"+imUrl;
                             var imageURL=imUrl;//"http://res.cloudinary.com/demo/image/fetch/"+imUrl;
                             for (var j=0; j<values.length; j++)
                             {
                                   var vari=values[j][6];
                                   var varis=vari.split("|")
                                   
                                   for (var k=0; k<varis.length; k++)
                                   {
                                       if(varis[k]==color && allImages[j][0]=="") //no existing image
                                       {
                                           allImages[j][0]=imageURL.replace('_640x640.jpg','').replace('_50x50.jpg','');
                                       
                                       }
                                   }
                                   
                             }
                             

                    }
  
  sheet.getRange(rng.getRow(), 20, allImages.length, 1).setValues(allImages)

}








