
function onEditAmTitle()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
    
      var rng=ss.getActiveRange();
      var row=rng.getRow();
      rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
      var values=rng.getValues();
      var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
      var ddValues=ss.getSheetByName("Dropdown Options").getDataRange().getValues();
      
      for (var i=0; i<values.length; i++)
      {     
            var vals=[values[i]]; //send one row but as 2D
            if(vals[0][12-1]==""){continue;}
            var col=12;
            onEditAmTitleEachRow(vals, row+i, col, mapValues, ddValues)
      
      }
      

}





function onEditAmTitleEachRow(vals, row, col, mapValues, ddValues) 
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var sheet=ss.getActiveSheet();
       

        if(vals==undefined)
        {
                     var rng=ss.getActiveRange();
                     var row=rng.getRow();
                     var col=rng.getColumn();
                     rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
                     var values=rng.getValues();
                     var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
                     var ddValues=ss.getSheetByName("Dropdown Options").getDataRange().getValues();
                     vals=values[0];
        
        
        }

        
        if(isLister(sheet))  //when someone is editing amazon title
        {
              var amTitle=vals[0][12-1].toString();
              if(amTitle==""){return 0};
              
              var returns=myLookupColor(replaceAll(amTitle,"_"," "), mapValues, 1, 27, 29)  //macth with column A
              var colors=returns[0];
              Logger.log(colors)
              var sizes=returns[1];
              var patterns=returns[2];
               if(patterns.length<1)
                  {
                       var searchTerms='No known pattern found, please make Terms manually';
                  }
               else
               {
                    var searchTerms=   getSearchTerms(patterns, colors, replaceAll(amTitle,"_"," "), sizes, mapValues) 
               }
               sheet.getRange(row, 19).setValue(searchTerms)

             //find material 
              var rowMaterial= myLookupFullWord(replaceAll(amTitle,"_"," "), mapValues, 11)
              if(rowMaterial!=null)
               { 
                  var material=mapValues[rowMaterial-1][11-1+1];
                  sheet.getRange(row, 23).setValue(material);
               }
              
              //create keyword sentence
              
              var category=findMyCategory(replaceAll(amTitle,"_"," "), mapValues);
              
              if(category!=null)
              {
                  var baseLine= 'Beautiful CCC KKK features a PPP pattern and design.';
                  
                  var pattern=mapValues[patterns[0]-1][29-1]; //patterns[0] is the row number of firt pattern  
                  var kwLine=baseLine.replace('CCC', colors[0].toLowerCase()).replace('KKK', category).replace('PPP', pattern.toLowerCase());
                 //Browser.msgBox(sheet.getRange(row, 18).getFormula());
                 if(sheet.getRange(row, 18).getFormula()!="" || sheet.getRange(row, 18).getValue()=="")    // only overwrite if ther is a formula or a value
                  {
                        //sheet.getRange(row, 18).setValue(kwLine)
                  
                  };
                  
                  
   
              }//category is not null
             
              
              
                    var matchedRow=myLookup(amTitle, mapValues, 1)
                    var color="";
                    if(matchedRow!=null)
                    {
                      color=mapValues[matchedRow-1][2-1];
                      
                    }
                    sheet.getRange(row, 21).setValue(color)
                    
                    var matchedRow2=myLookup(amTitle, mapValues, 4)
                    
                    var size="";
                    if(matchedRow2!=null)
                    {
                      size=mapValues[matchedRow2-1][5-1];
                      
                    }
                    sheet.getRange(row, 22).setValue(size);
                    
                   
                   var matchedRow3=myLookup(amTitle + vals[0][0], mapValues, 7)
                    var mycatagory="";
                    if(matchedRow3!=null)
                    {
                      mycatagory=mapValues[matchedRow3-1][8-1];
                      
                    }
                    sheet.getRange(row, 14).setValue(mycatagory)

              
              
              
              
              
              

        }// when column 12 is edited
       

  
}








































function manualValidity()
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        
        var startRow=rng.getRow();
        var lastRow=rng.getLastRow();
        var msg="<br>";
        
       for(var i=startRow; i<=lastRow; i++)
       {
        var row=i;//rng.getRow();
        var col=rng.getColumn();
        
        var rowValues=sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
        
        var amTitle=rowValues[0][12-1];
        if(amTitle=="")
        {
          continue
        }

        var amTitleN=amTitle.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
        var returns=myLookupColor(replaceAll(amTitle,"_"," "), mapValues, 1, 27, 29)  //macth with column A
              var colors=returns[0];
              var sizes=returns[1];
              var patterns=returns[2];
              
        var category=findMyCategory(amTitle, mapValues);
        
        var cRow=row;
        msg= msg+ '<br><br><b>Row '+cRow+":</b> " +checkValidity(rowValues, mapValues, sizes, colors, patterns, category, cRow);
       }
       
       
         var html = HtmlService.createHtmlOutput(msg)
                               .setTitle('Error Check Results')
                               .setWidth(300);
          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                        .showSidebar(html);

}


function checkValidity(rowValues, mapValues, sizes, colors, patterns, category, cRow)
   {
           // normalizing using https://stackoverflow.com/questions/20864893/javascript-replace-all-non-alpha-numeric-characters-new-lines-and-multiple-whi
           var ss=SpreadsheetApp.getActiveSpreadsheet();
           
           var variation=rowValues[0][7-1];
           var sourceTitle=rowValues[0][1-1];
           var amTitle=rowValues[0][12-1];
           var msg="";
           
           var variationN="";
           
           
           
           if(variation==""){variation=sourceTitle;} //if variation column is empty try to get it from OS title
           if(variation!="")
             {
                    variationN=variation.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
                    variationN=variationN.replace("|"," ");
               
             }
               
           
           var amTitleN=amTitle.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
           
           var mapValues2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mapping2').getDataRange().getValues();   
           var matchedSize="";
           var titleSize="";
          


           for(var i=12; i<mapValues2.length; i++)
           {
                       var sizeTemp=(mapValues2[i][4-1]).toString().toLowerCase();
                       if(sizeTemp==""){ continue; }
                       
                       
                      
                       if(fullWordIndexOf(variationN, sizeTemp)>=0)
                       {
                            matchedSize=sizeTemp;
                            titleSize=mapValues2[i][5-1];  //what should be in AM Title column E
                            break;
                       }
   
           
           
           }







           if(matchedSize=="")
           {
                    msg+='<br>-<font color="red"> We are unable determine size of the prouct on from OS/WM Name or Variation for row:  '+cRow +'</font>';
                     
          
            }
           
           
           
             //we have found a matched size
           else
           {
                       if(fullWordIndexOf(amTitleN, titleSize)<0)  // macthed size is not in variation
                       {
                            
                            msg+='<br>-<font color="red">'+ titleSize+' not found in AM Title</font>';
                            
                           // Browser.msgBox('"'+titleSize+ '" not found in Amazon Title');
                           // return 0;
                       
                       }
                       
                       else
                       {
                            //now find the matched size in AM title
                                           var matchedSize2="";
                                           for(var i=0; i<mapValues2.length; i++)
                                           {
                                                       var sizeTemp=mapValues2[i][5-1];
                                                       if(sizeTemp==""){ continue; }
                                                       
                                                       
                                                       
                                                       if(fullWordIndexOf(amTitleN, sizeTemp)>=0)
                                                       {
                                                            matchedSize2=sizeTemp;
                                                            break;
                                                       }
                                           }  
                                           
                                           if(titleSize==matchedSize2)
                                           { 
                                               msg+='<br>-<font color="green">'+ titleSize+' matched</font>';
                                               //msg=msg+"\n"+titleSize+ ' mathced'; 
                                               //Browser.msgBox(msg);
                                           }
                                           else {
                                           
                                                     { 
                                                           msg+='<br>-<font color="red">'+ titleSize+' did not match with '+ matchedSize2 +'</font>';
                                                   
                                                     }//Browser.msgBox(titleSize +" and "+ matchedSize2+ ' not mathced for row'+ cRow);}
                                                }
                                                
                       } //end of else
                 
           }
           
         // ----------------------***------------------
           //column L to S 
             
           
             var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
             var row=sheet.getActiveRange().getRow();
             for (var i=12-1; i<19-1; i++)
             {
                   var temp=rowValues[0][i];
                   temp=temp.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
                   //checks if duve cover missing
                   if(temp.indexOf('duvet')>=0 && temp.indexOf('duvet cover')<0)
                   {
                     
                           msg+='<br>-<font color="red">Duvet Cover missing in cell '+ sheet.getRange(row, i+1).getA1Notation() +'</font>';
                     //Browser.msgBox('Duvet Cover missing in cell '+ sheet.getRange(row, i+1).getA1Notation());
                     //return 0
                   
                   }
                   
                   
                   if(temp.indexOf('1 piece')>=0 || temp.indexOf('1 pc')>=0 || temp.toLowerCase().indexOf('single')>=0)
                   {
                       if(temp.indexOf("set")>=0)
                       {
                           
                           msg+='<br>-<font color="red">Set detected with one piece item</font>';
                           //Browser.msgBox('Set detected with one piece item for '+ cRow);
                           //return 0;
                       
                       }
                   
                   
                   }
                   
             
             }
            
             
             
             // https://stackoverflow.com/questions/1183903/regex-using-javascript-to-return-just-numbers
             var setIncludes=(rowValues[0][16-1]).toString().trim().toLowerCase();  //remove all spaces
             msg+=countIncludes(setIncludes, amTitle, mapValues, mapValues2);  //check validaity in set includes
             
             
             //----check category------------
            var catTitle= findMyCategory(amTitle, mapValues);
            var catOs=findMyCategory(rowValues[0][0],mapValues);
            if(catOs!=catTitle)
            {
                  if(catOs==null)
                  {
                      if(catTitle.indexOf('Bed in a Bag')>=0)
                      {
                          catOs="Comforter"
                      
                      }
                  
                  }
                  
                  
                  if(catOs!=null) 
                  {
                    
                    
                          msg+='<br>-<font color="red">'+ catOs+' did not match with '+ catTitle +'in AM Title</font>';
                    
                        //Browser.msgBox('"'+catOs+ '" found from source title didn\'t match '+catTitle+ ' found from Amazon Title for row '+cRow);
                  }
            
            }
            
            
            var catInc=findMyCategory(rowValues[0][15],mapValues);
            if(catInc!=catTitle)
            {
                  
                  
                  msg+='<br>-<font color="red">'+ catInc+' in Set Includes did not match with '+ catTitle +'in AM Title</font>';
                  //Browser.msgBox('"'+catInc+ '" found from Set Includes didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                 
            
            }
            
            
            
            var catDim=findMyCategory(rowValues[0][16],mapValues);
            if(catDim!=catTitle)
            {
                  
                  msg+='<br>-<font color="red">'+ catDim+' in Dimensions did not match with '+ catTitle +'in AM Title</font>';
                  //Browser.msgBox('"'+catDim+ '" found from Dimensions didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
   
            var catTerms=findMyCategory(rowValues[0][18],mapValues);
            if(catTerms!=catTitle && catTerms!=null)
            {
                  
                  msg+='<br>-<font color="red">'+ catTerms+' in Terms did not match with '+ catTitle +'in AM Title</font>';
                  //Browser.msgBox('"'+catTerms+ '" found from Search Terms didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
    
           
            var catColN=findMyCategory(replaceAll(rowValues[0][14-1],"-"," "),mapValues);
             if(catColN!=catTitle)
            {
                   msg+='<br>-<font color="red">'+ catColN+' in column N did not match with '+ catTitle +'in AM Title</font>';
                  //Browser.msgBox('"'+catColN+ '" found from column N didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
            
            //term size checking
            
            var sizeTitle =sizes[0]; //size determined from title for search terms
            var terms=rowValues[0][19-1]; //col S
            for (var i=0; i<mapValues.length; i++)
            {
                  var tempSize=mapValues[i][28-1];
                  if(terms.indexOf(tempSize)>=0)  //this size is found in terms
                  {
                        if(tempSize!=sizeTitle && tempSize!="")
                        {     
                              
                              
                               msg+='<br>-<font color="red">Invalid size '+ sizeTemp+' found in serch term.</font>';
                              //Browser.msgBox('Invalid size '+sizeTemp + ' found in serch term for row '+ cRow);
                              //return 0;
                        
                        }
                  
                  }
            
            
            
            }
             
            return msg
   
            
   
   }





  function countIncludes(setIncludes, amTitle, mapValues, mapValues2)
  {
      
      var msg=""
      setIncludes=replaceAll(setIncludes, "_", "");  //eliminate under score
      
      setIncludes=replaceAll(setIncludes, "( ","(");  //elimiate any comma after the bracket
      setIncludes=replaceAll(setIncludes, " )",")");
       
       
      amTitle=replaceAll(amTitle.toLowerCase(), "_", " ");
      
      var setPlural=setIncludes;  //set icnludes for plural detection
      
      var count=0;
      
      for( var r=1; r<15 ; r++)  //random loop assuimg max 15 items in a set
      {      
      
            for (var i=1; i<mapValues2.length; i++)
            {
                    var temp=(mapValues2[i][7]).toString().toLowerCase().trim();
                    if(setIncludes.indexOf(temp)>=0 && temp!="")
                    {     
                          var tempCount=mapValues2[i][8];  //col H
                          Logger.log(temp)
                          
                          count=count+Number(tempCount);
                          setPlural=replaceAll(setPlural,temp,"|"+temp);  // add a seperator for later use to find plural error
                          Logger.log(setPlural)
                          setIncludes=setIncludes.replace(temp, "");  //remove the matched count so it does not get counted again
                         
                          break;
                    }
            
            
            }
       }     
      if(count==0){
         msg+='<br>-<font color="red">No Set Includes found.</font>';
      
      };
      
      
      var titlePieces="";
      for (var i=0; i<mapValues.length; i++)
      {
                var tempValue=(mapValues[i][17-1]).toLowerCase();//
                
                if(tempValue !="" && fullWordIndexOf(amTitle,tempValue)>=0)
                {
                        titlePieces=tempValue;
                        break;
                
                }
      
      }
      
      
      
      if(titlePieces=="")
      {
             msg+='<br>-<font color="red">No piece found in AM Title.</font>';
      
      }
      var titleCount=Number(titlePieces.match(/\d+/g)[0]);  //https://stackoverflow.com/questions/1183903/regex-using-javascript-to-return-just-numbers
     
      if(count>1 || titleCount>1)
      {
              if(amTitle.toLowerCase().indexOf("set")<0)
              {
                  msg+='<br>-<font color="red">Set not found with multiple peice item</font>';
              }
      
      }
     
     
      if(count==titleCount)
      {
           
           msg+='<br>-<font color="green">'+count+' piece matched</font>';
           // Browser.msgBox(count +' piece matched');
           // return 0;
      }
      
      else
      {
           msg+='<br>-<font color="red">Set Includes count '+count+' not matched with Title piece count- '+titleCount+'</font>';
          //Browser.msgBox("Set includes count: "+count+ " didn't match with title piece count "+titleCount);
          //return 0;
      
      }
      
      
      //plural error finding start
      
      var setPluralArr=setPlural.split("|");
      
          for (var i=0; i<setPluralArr.length; i++)
          {
              var set=setPluralArr[i].toString().toLowerCase();
              
              
              if(set.indexOf('one')>=0 || set.indexOf('1')>=0)
             {
             
                    set=replaceAll(set," ","");
                    var lastChar=set.slice(-1);
                    if(lastChar=='s')
                    {
                        msg+='<br>-<font color="red">Suspected plural noun used with 1 piece</font>';
                  
                    }
              }
              
          
          }
      
      
      
      
      
      
      
      
      
      
      return msg;
      
      
      
      
      
  
  
  }



















function findMyCategory(amTitle, mapValues)
{
     for(var i=30-1; i<mapValues[0].length; i++)
     {
           var phrase=(mapValues[0][i]).toString().toLowerCase();
           if(fullWordIndexOf(amTitle, phrase)>=0)
           {
               return phrase;
           
           }
           
           
     }

    return null


}









function myLookupColor(val, mapVals, col1, col2, col3)
{
     val=val.toLowerCase();
     var colors=[];
     var sizes=[];
     var patterns=[];
    // var valArr=replaceAll(val,',','').toLowerCase().split(' ');
     var index=10;
     for (var i=1; i<mapVals.length; i++)
     {
          
           //if(mapVals[i][col1-1]==""){continue};
           
           //color
           var tempColor=mapVals[i][col1-1];
           var nn1=fullWordIndexOf(val, mapVals[i][col1-1]);
           if(nn1>=0 && tempColor!='')
           {
                
                if(mapVals[i][col1]!="")
                    {colors.push([nn1,(mapVals[i][col1-1])])}; //store the color in variation
                if(mapVals[i][col1]!="")                
                    {colors.push([nn1,(mapVals[i][col1])])};  //push the color that matches with the mapped base color
                   //return i+1;
            } 
            
            
            //size
            if(fullWordIndexOf(val, mapVals[i][col2-1])>=0)
           {
                
                if(mapVals[i][col2]!="")
                    {sizes.push((mapVals[i][col2]))}; //store the size in variation
                
            } 
            
            
            //
            //pattern
            var multiplePatterns=mapVals[i][col3-1];   
            var patternArr=multiplePatterns.split(','); // we can enter comma seperated patterns in mapping 
            
            for( var p=0; p<patternArr.length; p++)
            {
                   var indx=fullWordIndexOf(val, patternArr[p].trim())
                   if(indx>=0)
                   {
                        
                        
                        if(mapVals[i][col3-1]!="")
                            {
                                            patterns.push([indx, i+1]); //push the row in patterns
                            }
                        
                    } 
            }
            
            
            
            
     } 
     
      //http://stackoverflow.com/questions/16096872/how-to-sort-2-dimensional-array-by-column-value
      patterns.sort(sortFunction);
      colors.sort(sortFunction);
      
      function sortFunction(a, b) {
          if (a[0] === b[0]) {
              return 0;
          }
          else {
              return (a[0] < b[0]) ? -1 : 1;
          }
      }
      
      
     var patternsTemp=[];
    
     for (var p=0; p<patterns.length; p++)
     {
           
           patternsTemp.push(patterns[p][1]);
     
     }
     
     
       
     var colorsTemp=[];
    
     for (var c=0; c<colors.length; c++)
     {
           
           colorsTemp.push(colors[c][1]);
     
     }
     
     
     
     
     
     
     
     var a=colorsTemp.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     var b=sizes.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     var c=patternsTemp.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     
     Logger.log(colorsTemp)
     if(patternsTemp.length>3) {patternsTemp=patternsTemp.slice(0,3)}; //limit maximum 3 patterns
     return [a,b,c];  
}






function getSearchTerms(patterns, colors, amTitle, sizes, mapValues) 
{
     
     var terms=[];
     var size=sizes[0];  //just one size
     
     for(var i=30-1; i<mapValues[0].length; i++)
     {
           
           if(amTitle.toLowerCase().indexOf(mapValues[0][i].toLowerCase())<0){continue;}
           
           for (var j=0; j<patterns.length; j++)
           {
                      var patternRow=patterns[j];
                      var sTerm=mapValues[patternRow-1][i];  // this is the base search term
                      
                        for (var k=0; k<colors.length; k++)
                        {
                                    var sTerm1=replaceAll(sTerm, "CCC", colors[k]);
                                    sTerm1=replaceAll(sTerm1, "SSS", size);
                                    terms.push(sTerm1); 
                      
                        }
                      
                      
           
           }
           
       
     }
  
 
  
  
  var searchTerms= terms.join(", ").split(", "); 
  var uniqTerms= searchTerms.filter(function(item, i, ar){ return ar.indexOf(item) === i; })
  
  var ret=uniqTerms[0]; //uniqTerms.join(", ");
  
     for (var i=1; i<uniqTerms.length; i++)
     {
           if(ret.length<850)
           {
               ret=ret+", "+uniqTerms[i];
           
           }
     
     }
  
    return ret.toLowerCase();
  
  
  
}




function lookup5(l_value, sheet2, lookup_col, pick_up_col, value_or_row) {
  var last_row2 = sheet2.getLastRow();
  
  if (last_row2<2) { last_row2 = 2; }
  
  var ar=sheet2.getRange(2,lookup_col,last_row2-1, pick_up_col-lookup_col+1).getValues();
  
  var flag=0;
  for (var i=0; i<last_row2-1; i++)
  {
    var temp1 = ar[i];
    var temp=temp1[0];
    if(temp==l_value)
    {
      flag=1;
      if (value_or_row=="value")   
      {
        return ar[i][temp1.length-1];
        break;
      }
      
      if (value_or_row=="row") {
        return i+2;
        break;
      }
    }
  }
  if(flag==0){ return null; }
}




























//These scripts will help to show already listed images in sidebar for reviewing







function re_checkImagesByVariation()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();
    
     var sbHtml='<br>';
     var htmlArr=[];
                    
    for (var i=0; i<values.length; i++)
    {
          var variation=values[i][7-1];
          var img1=values[i][20-1];
          var img2=values[i][24-1];
          var img3=values[i][25-1];
          var img4=values[i][26-1];
          
          var type1="Regular";
          if(img1.indexOf('a_hflip')>=0)
          {
              type1="Flipped";
          }
          
          if(img1.indexOf('c_crop')>=0)
          {
              type1="Cropped";
          
          }
          
          
           var type2="Regular";
          if(img2.indexOf('a_hflip')>=0)
          {
              type2="Flipped";
          }
          
          if(img2.indexOf('c_crop')>=0)
          {
              type2="Cropped";
          
          }
          
          var type3="Regular";
          if(img3.indexOf('a_hflip')>=0)
          {
              type3="Flipped";
          }
          
          if(img3.indexOf('c_crop')>=0)
          {
              type3="Cropped";
          
          }
          
          
          var type4="Regular";
          if(img4.indexOf('a_hflip')>=0)
          {
              type4="Flipped";
          }
          
          if(img4.indexOf('c_crop')>=0)
          {
              type4="Cropped";
          
          }
          img1=img1.replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          img2=img2.replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          img3=img3.replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          img4=img4.replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          var tempHtml='<br><br>Variation: <b>'+variation+ '</b><br><br>Main Image: '+type1+'<br><img src="'+img1+'" alt="Not Available" style="width:auto; height:200px;"><br>'          
                                                    + '<br>Image 2 '+type2+'<br><br><img src="'+img2+'" alt="Not Available" style="width:auto; height:200px;"><br>'            
                                                    + '<br>Image 3 '+type3+'<br><br><img src="'+img3+'" alt="Not Available" style="width:auto; height:200px;"><br>'
                                                    + '<br>Image 4 '+type4+'<br><br><img src="'+img4+'" alt="Not Available" style="width:auto; height:200px;"><br>';
                                                    
          htmlArr.push(tempHtml);                                           
    
    }
    
    
    
        
        
                  
                   
                   
                   
      sbHtml=sbHtml+htmlArr.join('<br>');
      Logger.log(sbHtml)
          
                        
                        
      if(sbHtml!='<br>')
      {
            var imhtml = HtmlService.createHtmlOutput(sbHtml)
            .setTitle('Images')
            .setWidth(300);
            SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
            .showSidebar(imhtml);
        
      }



    
    
    
    
    
    
    
 }













function re_checkImagesByImagePosition()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();
    
     var sbHtml='';
     var htmlArr=[];
     
  for (var k=1; k<=4; k++)
  {
          for (var i=0; i<values.length; i++)
          {
                var variation=values[i][7-1];
                
                if(k==1)
                {
                    var img=values[i][20-1];
                }
                else
                {
                    var img=values[i][24+k-2-1];
                
                }
                
                
                var type="Regular";
                if(img.indexOf('a_hflip')>=0)
                {
                    type="Flipped";
                }
                
                if(img.indexOf('c_crop')>=0)
                {
                    type="Cropped";
                
                }
                
                

                  var position="Main Image";
                
                
                 if(k>1)
                {
                  position="Image "+k;
                }
                
                
                img=img.replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
                
                var tempHtml='Variation: <b>'+variation+ '</b><br>'+position+': '+type+'<br><img src="'+img+'" alt="Not Available" style="width:auto; height:200px;"><br>'          
                
                htmlArr.push(tempHtml);                                           
          
            }
            
            sbHtml=sbHtml+"<br><br><br>"+htmlArr.join('<br>')+'-------------------------------------------<br>-------------------------------------------';
            htmlArr=[];

  }  
    
    
        
        
                  
                   
                   
                   
          
                        
                        
      if(sbHtml!='<br>')
      {
            var imhtml = HtmlService.createHtmlOutput(sbHtml)
            .setTitle('Images')
            .setWidth(300);
            SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
            .showSidebar(imhtml);
        
      }



    
    
    
    
    
    
    
 }






function checkPrimaryImages()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();
    
     var sbHtml='<br>';
     var htmlArr=[];
                    
    for (var i=0; i<values.length; i++)
    {
          var variation=values[i][7-1];
          var img1=(values[i][20-1]).replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          var img2=(values[i][24-1]).replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com');
          var img3=(values[i][25-1]).replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com')
          var img4=(values[i][26-1]).replace('https://www.dropbox.com','https://www.dl.dropboxusercontent.com')
          

          
          var tempHtml='<br><br>Variation: <b>'+variation+ '</b><br><img src="'+img1+'" alt="Not Available" style="width:auto; height:200px;"><br>Row:   '+(row+i)+'</br>'          
                                                    
          htmlArr.push(tempHtml);                                           
    
    }
    
    
    
        
        
                  
                   
                   
                   
      sbHtml=sbHtml+htmlArr.join('<br>');
      Logger.log(sbHtml)
          
                        
                        
      if(sbHtml!='<br>')
      {
            var imhtml = HtmlService.createHtmlOutput(sbHtml)
            .setTitle('Images')
            .setWidth(300);
            SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
            .showSidebar(imhtml);
        
      }



    
    
    
    
    
    
    
 }
 
 
 
 
 
 
 
 
 
 function checkImageVsVariables()
 {
     var ss=SpreadsheetApp.getActiveSpreadsheet();
     var sheet=ss.getActiveSheet();
     var rng=sheet.getActiveRange();
     
     var row=rng.getRow();
     var lr=rng.getLastRow();
     
     var values=rng.getValues();
     var images=[];
     var indxs=[];
     
     for (var i=0; i<values.length; i++){
     
       var img=(values[i][7-1]).split("|")[0];
       if(images.indexOf(img)<0){
         images.push(img);
         indxs.push(i);
       }
       
       
     }
     
     var html="";
     for(var i=0; i<indxs.length; i++){
         
         html+='<br><br><br>'
         html+=images[i];
         html+='<br>'
         html+='<img src="'+values[indxs[i]][20-1]+'" alt="Not Available" style="width:auto; height:200px;">'
         var variables=[];//[values[indxs[i]][35-1], values[indxs[i]][36-1], values[indxs[i]][37-1], values[indxs[i]][38-1],
         for (var j=35; j<=50; j++){
           if(values[indxs[i]][j-1]!=""){
               variables.push(values[indxs[i]][j-1]);
           }
         }//end of for
         html+='<br><br>'+variables.join(" || ")  
     }
     
     
     




     var imhtml = HtmlService.createHtmlOutput(html)
     .setTitle('Images')
     .setWidth(300);
     SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .showSidebar(imhtml);
     
     
 
   var a=10
 
 }
 







