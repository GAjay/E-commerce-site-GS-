//Main Function 
function onFromSubmit(e){
          var id = e.values[1];//take value of spreadsheet of 2 col latest
          
          var cat = e.values[2];//take value of spreadsheet of 3 col latest
          
          var emailaddress = e.values[3];//take value of spreadsheet of 4 col latest 
                
          var k=0;
          
          var idnum  ;
          
          var bool = idchk(id);
          
          if(bool === true){
          
             switch(cat){//starting of switch case
                         case 'MOBILE'://choose category
                                   idnum="1cXsTljmCR8tIcW8sU6bECFP6agzVviDIr7r5TnmunPI";
                                  SendMailtouser(idnum,k,id,emailaddress,cat);
                           break;//break the case of mobile
          
                           case 'BOOK':
                                   //choose category
                                  
                                  //GET Book SpreadSheet using spreadsheet id function
                                   idnum="1pYITJ7Ujxuo6nhQGWa5exaRLsByJ-6Hvpun-yCA0uHg";
                                   SendMailtouser(idnum,k,id,emailaddress,cat);
                                   
                           break;//break the case of Book
          
                           case 'LAPTOP':
                                   //choose category
                                  
                                  //GET Laptop SpreadSheet using spreadsheet id function
                                   idnum ="1IXBdZbdDEn-B5k__UUwzqf8efE16XtABw-NIJhLY0Po";
                                   SendMailtouser(idnum,k,id,emailaddress,cat);
                                   
                           break;//break the case of laptop
              default:
               
               var message = "choose again";
               
               GmailApp.sendEmail(emailaddress,"Shipping details",message);
           }
          }
           else //id length is not valid
           {
           var message = "choose 5 digit id";
           GmailApp.sendEmail(emailaddress,"Shipping details",message);
           }

}//end of onFrom Submit

//function of check id length 
function idchk (id){
                    var inode=id;
                    if(inode.toString().length ===5)
                    {
                     return true;
                    }
                    else
                    {
                      return false;
                    }
                  }
                  
  //closure Function for get value of notation                
  var createNote = function(name) {
  
  var num;
   
   return {
           getNum: function() {
                                return num;
                              },
           
           setNum: function(newNum){
            
                    if(newNum == 1){
                                    num = 2;
                                   }
                    
                    else if(newNum ==2){num = 3;}
                    
                    else if(newNum ==3){num =4;}
    }
  }
}


function SendMailtouser(idnum,k,id,emailaddress,cat){
           var key=idnum;
           
           var k=k;
           
           var id=id;
           
           var emailaddress= emailaddress;
           
           var cat=cat;
           
           var ss = SpreadsheetApp.openById(key);//open mobile sheet using its id 
                                  
           var sheet = ss.getSheets()[0];//select sheet 0 means 1 first sheet  
           //var startRow = 2;//
          //var numRows = 4;//number of cloum
         // Fetch the range of cells A2:D4
          var dataRange = ss.getRangeByName("A2:D4");//Get range from the notations
                                  
          var data = dataRange.getValues();//get values of spreadsheet according to datarange fn
                                  
          var boolean = false;
                                  
          //take data from mobile spreadsheet
          
          for(var i=0;i<data.length;i++){//for  loop for data 
                                         var row = data[i];//array calling
                                                  
                                         var id1 = row[0];//id value
                                                  
                                         k++;//counter
                                                  
                                         if(id == id1)//checking id exist or not
                                                  {
                                                    var stock = row[3];//take stock data
                                                    
                                                    var company_name =row[2];
                                                    
                                                    var author_name=row[1];
                                                    
                                                    if(stock > 0)//checking stocks value 
                                                    {
                                                     //message = "YOUR"+""+cat+""+"IS SHIPPED"+stock+"  "+k;//message
                                                     
                                                       var template = HtmlService.createTemplateFromFile('mail');
                                                     //variabes of mail.html file
                                                       template.main ="Mobile Report"//main variable of mail.html file 
                                                     
                                                       template.category= cat.toLowerCase();
                                                     
                                                       template.ID="ID";
                                                     
                                                       template.Name="Company Name";
                                                     
                                                       template.os="Operating System";
                                                     
                                                       template.name = company_name;
                                                     
                                                       template.author = author_name;
                                                     
                                                       template.id = id;
                                                      //template.scriptUrl = scriptUrl;
                                                    //template.serialNumber = getGUID();  // Generate serial number for this response
                                                      var html = template.evaluate().getContent();//recipient = Session.getActiveUser().getEmail();  

                                                      GmailApp.sendEmail(emailaddress, "Shipping Details", 'Requires HTML', {htmlBody:html} );//message sending to form emailaddres
                                                     
                                                      stock--;//decremeant in stock variable
                                                     
                                                      var note =createNote(k);//calling coluser function
                                                     
                                                      note.setNum(k);//calling colsuer function internal method setNum()
                                                     
                                                      var nota = note.getNum();//calling colsuer function internal method getNum()
                                                     
                                                      var range = sheet.getRange (nota, 4);
                                                     // Logs "A1:E2"
                                                     
                                                      var s = range.getA1Notation();//get string value of curren column
                                                     //Logger.log(s);//can be removed
                                                     
                                                      ss.getRange(s).setValue(stock);
                                                     //Logger.log(s);//can be removed
                                                     //ss.getRange(s).setValue(stock);
                                                    }
                                                   
                                                   else //if stock is not avaiable
                                                       {
                                                        GmailApp.sendEmail(emailaddress, "shipping details", " transaction failed");//if stock is not avaible
                                                       }
                                                  return boolean =true;//id exist 
                                                       
                                                  break;//break for "for in"loop
                                                }//end of first if of "for in" loop
                                             }//end of "for" loop
                                             
                                 if(boolean != true)//checking boolean variable status 
                                                   {
                                                    message="transaction failed"+"please enter valid id "+id;//message for invalid id
                                                    
                                                    GmailApp.sendEmail(emailaddress,'shipping details',message);
                                                   }
}
