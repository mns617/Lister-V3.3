

//shows overstock images, 
//calls WM image show if url has walmart line 18

function imShowSideBar() {

    
        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
        if(!(isLister(sheet))){return 0}
        if(col!=8){return 0}        
        
                    var headers = {                           
                        'ostkid': 'OSTK-VIP_18-A77359'                         
                      };                           
                      var option = {                            
                        "headers": headers,
                        'muteHttpExceptions' : true                           
                      };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    if(getURL.indexOf('walmart.com')>=0)
                    {
                            showWmImages();
                            return 0;
                            
                    
                    }
                    if(getURL.indexOf('walmart.com')>=0)
                    {
                            showAliImages()
                            return 0;
                            
                    
                    }
                    
                    if(getURL.indexOf('aliexpress.com')>=0)
                    {
                        showAliImages()
                    }
                    
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                    
                    var n1=html.indexOf('s-h-title');
                    var n2=html.indexOf("<",n1);
                    var title=html.slice(n1+11,n2-1); 
                    
                //    var folderId = "0Bw-TXeLyArDnLTBabkVxUXBoeVk";
                    
                  //  var tempFolder=DriveApp.getFolderById(folderId).createFolder(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMMM-yyyy hh:mm a")+"--"+title);; 
                   // var folderUrl=tempFolder.getUrl();
                    
                    var n1=html.lastIndexOf('<div class="container">');
                    var n2=html.indexOf('</ul>',n1)
                    var html2=html.slice(n1,n2)
                    var sbHtml='<br>';
                    var imgUrlArr=html2.split('data-max-img');
                    var retImgUrl="";
                        for (var j=1; j<imgUrlArr.length; j++)  //when there is variation, index 0 has garbage data
                        {
                           var longUrl=imgUrlArr[j];
                              var l1=longUrl.indexOf("ak1");
                              var l2=longUrl.indexOf(">",l1);
                              var imUrl=longUrl.slice(l1,l2-1);
                              
                              //var imageURL=(imUrl).replace("ostkcdn.com","ostkcdn.com.rsz.io")+"?flip=x"
                              
                              var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                              var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/http://"+imUrl;
                              var imageURL="http://res.cloudinary.com/demo/image/fetch/http://"+imUrl;
                              if(retImgUrl=="")
                              {retImgUrl=imageURL}
                              if(sMode=="on") //when sports ads are being made, past primary image
                              {
                                   sheet.getRange(sheet.getActiveRange().getRow(), 20).setValue(imageURL); return 0;
                              }
                              sbHtml=sbHtml+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'
    
                              +'<form>'
                                +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                                +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                                +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'
    
                               +'</form>'
                               +'<br><hr><br>'
                              
                              
                              
                              
                              
                              
                              
                              var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;

                          //var imBlob=UrlFetchApp.fetch(imageURL).getBlob();
                          
                          //var imFile=tempFolder.createFile(imBlob);
                          //imFile.setName(imPhrase+" "+ j+".jpg");
                        }
                        
                        
                        
                        
                        
                        
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }     
                        
                        
                        
                        
               return retImgUrl         
                        
                        
                        

  }
  
  
  
    
  
  

