function walmarQty(url, myVar) 
{
      
      
      
      var url="https://www.walmart.com/ip/Mainstays-Garden-Floral-Bed-in-a-Bag/337396440";
      myVar="Full";
         
      var option = {
                      'muteHttpExceptions' : true
          };

      var html = UrlFetchApp.fetch(url, option).getContentText();        
      var jsonData=getMyJson(html);
      
      
      
      
      
      
      var product=jsonData.product;
      var primaryProduct=product.primaryProduct; //varaition map starts with base product
      var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
            
      if(varMap==undefined)
      {
                  //no variation item
                  myVar="";
      }
            
            
            
       else
       {
            
            
            var flag1=0;
            var flag2=0;
            // these two arrays will all variation information
            var cv=varMap.actual_color;
            if(cv!=undefined)
              {var colorVars=cv.variants;}
            else
              {flag1=1;}
              
              
             var sv= varMap.size;
             if(sv!=undefined) 
              {var sizeVars=sv.variants;}
             else
             {flag2=1;}
            
      
        }
      
      var myProducts=jsonData.product.products;
      var arrP=[];
      var count=0;
      var arrTemp=[];
      
      var desiredVarId="";
      
      
      for (var i in myProducts)
      {
            var id=myProducts[i].usItemId;
            
            if(myVar==""){
            
                desiredVarId=id;
                break;    
            
            }; //when single variation
            
            var variantsProp=myProducts[i].variants; //variants of this product
            
            var count=0;
            var variation="";
            //get the variant details
            if(flag1==0 && flag2==0)
            {
              var sizeProp=variantsProp.size;
              var sizeName=sizeVars[sizeProp].name;
              
              var colorProp=variantsProp.actual_color;
              var colorName=colorVars[colorProp].name;
              
              var skugridVar=sizeName+'|'+colorName;
            }
            
            else if  (flag1==0)  //only color vari
            {
              var colorProp=variantsProp.actual_color;
              var colorName=colorVars[colorProp].name;
              
              var skugridVar=colorName;
            }
            
            
            else if  (flag2==0)  //only color vari
            {
              var sizeProp=variantsProp.size;
              var sizeName=sizeVars[sizeProp].name;
              var skugridVar=sizeName;
              
            }
            
            if(skugridVar==myVar)
            {
               desiredVarId=id
                break;    
            }
                      
            
      }         
      
      
      
      if(desiredVarId=="")
      {return "N/A"};
      
            var apiUrl='http://api.walmartlabs.com/v1/items?ids='+id+'&apiKey=2ry3zt8p73k5cggm4ytm7uc4';
            var data=UrlFetchApp.fetch(apiUrl).getContentText();
            var jsonData=JSON.parse(data);
      
           var myItems =jsonData.items;   
            for (var j in myItems)
            {
                  var myItem=myItems[j];
                  var myStock=myItem.stock;
                  
                  if(myStock=="Available")
                  {return 10}
                  else if(myStock=="Limited supply")
                  {return 4}
                  else
                  {
                    return 0
                  }
            }
      
        
       
      
      
      
}
