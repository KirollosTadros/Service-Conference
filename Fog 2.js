
  var app= SpreadsheetApp;
  var spreadSheet= app.getActiveSpreadsheet();
  var Sheet1=spreadSheet.getSheetByName("Form Responses 1");
  var Sheet2=spreadSheet.getSheetByName("IN");
  var Sheet3=spreadSheet.getSheetByName("Waiting");
  var Sheet4=spreadSheet.getSheetByName("Statistics");
  var Sheet5=spreadSheet.getSheetByName("Exchange");
  var Sheet6=spreadSheet.getSheetByName("Osar");
  var full="شامل كل حاجة (اشتراك كامل)";
  var half="شامل أكل فقط (نصف اشتراك)";
  var zero="بدون أكل أو اتوبيس أو سرير (مجانا)";
  var Exchange="تحويل من فوج اخر";
  var bus= "أتوبيس";
  var car="سيارة";
  var d1= "اثنين, ثلاثاء, أربعاء, خميس, جمعة";
  var d2= "اثنين, ثلاثاء, أربعاء, خميس";
  var d3= "اثنين, ثلاثاء, أربعاء"
  var d4= "اثنين, ثلاثاء"
  var d5= "ثلاثاء, أربعاء, خميس, جمعة"
  var d6= "ثلاثاء, أربعاء, خميس"
  var d7= "ثلاثاء, أربعاء"
  var d8= "أربعاء, خميس, جمعة"
  var d9= "أربعاء, خميس"
  var d10= "خميس, جمعة"
  var Servant= "خادم";
  var Female= "خادمة";
  var max=parseInt(Sheet4.getRange(1, 3).getValue());
  var children=parseInt(Sheet4.getRange(2, 2).getValue());
  var adults=parseInt(Sheet4.getRange(6, 2).getValue());
  stat();
  setDropdownMove();
  setDropdownChild();
  setDropdownDays();
  
 function setDropdownDays(){ 
 
  var list =[d1, d2, d3,d4,d5,d6,d7,d8,d9,d10];
   var dynamicList = Sheet2.getRange("K2:K"+Sheet2.getLastRow());
   var arrayValues = dynamicList.getValues();
   var rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("K2:K"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("K2:K"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("K2:K"+Sheet3.getLastRow()).setDataValidation(rangeRule);
}

function setDropdownMove(){   
 
  var list =[bus,car];
   var dynamicList = Sheet2.getRange("I2:I"+Sheet2.getLastRow());
   var arrayValues = dynamicList.getValues();
   var rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("I2:I"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("I2:I"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("I2:I"+Sheet3.getLastRow()).setDataValidation(rangeRule);
   
   dynamicList = Sheet2.getRange("J2:J"+Sheet2.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("J2:J"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("J2:J"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("J2:J"+Sheet3.getLastRow()).setDataValidation(rangeRule);
}

function setDropdownChild(){   
 
  var list =[full,half,zero, ""];
   var dynamicList = Sheet2.getRange("T2:T"+Sheet2.getLastRow());
   var arrayValues = dynamicList.getValues();
   var rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("T2:T"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("T2:T"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("T2:T"+Sheet3.getLastRow()).setDataValidation(rangeRule);
   
   dynamicList = Sheet2.getRange("W2:W"+Sheet2.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("W2:W"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("W2:W"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("W2:W"+Sheet3.getLastRow()).setDataValidation(rangeRule);
   
   dynamicList = Sheet2.getRange("Z2:Z"+Sheet2.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   // set the dropdown validation for the row
   Sheet2.getRange("Z2:Z"+Sheet2.getLastRow()).setDataValidation(rangeRule); // set range to your range
   dynamicList = Sheet3.getRange("Z2:Z"+Sheet3.getLastRow());
   arrayValues = dynamicList.getValues();
   rangeRule = app.newDataValidation().requireValueInList(list);
   Sheet3.getRange("Z2:Z"+Sheet3.getLastRow()).setDataValidation(rangeRule);
}
   
   //on Sheet edit to edit Expected and statistics
function onEdit(e) 
{
if(e.source.getActiveSheet().getName()=="IN")
{
var range=e.range;
var row=range.getRow();

if(row!=1)
{
 Sheet2.getRange(row, 2).setValue(calculate(Sheet2,row));  //editing expected value
  
 if(parseInt(Sheet2.getRange(row, 2).getValue())==parseInt(Sheet2.getRange(row, 3).getValue()))
 { Sheet2.getRange(row, 3).setBackgroundColor('#0FFF00');}   
else if((parseInt(Sheet2.getRange(row, 3).getValue())<parseInt(Sheet2.getRange(row, 2).getValue()))&&(Sheet2.getRange(row, 7).getValue()!=Exchange))  
{ Sheet2.getRange(row, 3).setBackgroundColor('#FF0000');}
  else if(parseInt(Sheet2.getRange(row, 3).getValue())>parseInt(Sheet2.getRange(row, 2).getValue()))
  {
    Sheet2.getRange(row, 3).setBackgroundColor('#FFFF00');
  }
  else if((parseInt(Sheet2.getRange(row, 3).getValue())<parseInt(Sheet2.getRange(row, 2).getValue()))&&(Sheet2.getRange(row, 7).getValue()==Exchange)) 
{ Sheet2.getRange(row, 3).setBackgroundColor('#A9A9A9');}
  
    else
    {   
      Sheet2.getRange(row, 3).setBackgroundColor('#FFFFFF');
  }
  //Statistcs edit
  Statistics();
  }

}


if(e.source.getActiveSheet().getName()=="Waiting")
{
var range=e.range;
var row=range.getRow();

if(row!=1)
{
 Sheet3.getRange(row, 2).setValue(calculate(Sheet3,row));  //editing expected value
  
 if(parseInt(Sheet3.getRange(row, 2).getValue())==parseInt(Sheet3.getRange(row, 3).getValue()))
 { Sheet3.getRange(row, 3).setBackgroundColor('#0FFF00');}   
 else if((parseInt(Sheet3.getRange(row, 3).getValue())<parseInt(Sheet3.getRange(row, 2).getValue()))&&(Sheet3.getRange(row, 7).getValue()!=Exchange))
  { Sheet3.getRange(row, 3).setBackgroundColor('#FF0000');}
  
  else if((parseInt(Sheet3.getRange(row, 3).getValue())<parseInt(Sheet3.getRange(row, 2).getValue()))&&(Sheet3.getRange(row, 7).getValue()==Exchange))
  { Sheet3.getRange(row, 3).setBackgroundColor('#A9A9A9');}
  else if(parseInt(Sheet3.getRange(row, 3).getValue())>parseInt(Sheet3.getRange(row, 2).getValue()))
  {
    Sheet3.getRange(row, 3).setBackgroundColor('#FFFF00');
  }
    else
    {   
      Sheet3.getRange(row, 3).setBackgroundColor('#FFFFFF');
  }
  //Statistcs edit
  Statistics();
  }
}

var the_range = e.range;
if((the_range.getRow()==1)&& (the_range.getColumn()==3) &&(e.source.getActiveSheet().getName()=="Statistics"))
{
  if(parseInt(Sheet4.getRange(1,3).getValue())-parseInt(e.oldValue)>5)
  {
    Sheet4.getRange(1,3).setValue(e.oldValue);
    var name=Browser.msgBox('Large amount', 'Please dont increase max fog number more than 5 at a time', Browser.Buttons.OK);
  }
  if (parseInt(Sheet4.getRange(1,3).getValue())<parseInt(e.oldValue))
  {
      var name=Browser.msgBox('Decreasing', 'Decreasing Max fog amount doesnt sent from IN to waiting you should change it by your self', Browser.Buttons.OK);

  }
        var name=Browser.msgBox('Saved', 'New Max Fog amount saved successfully', Browser.Buttons.OK);

}
  
  //Sending from waiting
   Statistics();
    var flag=0;
    while(Sheet3.getLastRow()>1)      //Check if there is someone waiting
  { 
    if(adults+children<max)        //Make Sure there is an empty place in the IN sheet
    {
      flag=1;
    if((max-adults-children==1)&&( Sheet3.getRange(2, 20).getValue()==full||Sheet3.getRange(2, 23).getValue()==full||Sheet3.getRange(2, 26).getValue()==full))
   {break;}
   else
   {
   var source_range = Sheet3.getRange("A"+2+":AD"+2);
    var target_range = Sheet2.getRange("A"+(Sheet2.getLastRow()+1)+":AD"+(Sheet2.getLastRow()+1));
  source_range.moveTo(target_range);
  //Sheet3.getRange("A3:AD").moveTo(Sheet3.getRange("A2:AD"));
     Sheet3.deleteRow(2);
    Statistics();
  }
    }
    else
    {
    break;
    }
  }
    if(flag==1)
  var name=Browser.msgBox('Done', 'Waiting insertion is done', Browser.Buttons.OK);

   Statistics();
   
   //end send from waiting

}


//This method handles when a row is deleted to allow the waiting to be IN
function onChange(e)
{
  if(e.changeType=="REMOVE_ROW"||e.changeType=="REMOVE_GRID")
  {
 Statistics();
 var flag=0;
    while(Sheet3.getLastRow()>1)      //Check if there is someone waiting
  { 
    if(adults+children<max)        //Make Sure there is an empty place in the IN sheet
    {
      flag=1;
    if((max-adults-children==1)&&( Sheet3.getRange(2, 20).getValue()==full||Sheet3.getRange(2, 23).getValue()==full||Sheet3.getRange(2, 26).getValue()==full))
   {break;}
   else
   {
      Sheet3.getRange(Sheet3.getLastRow(), 32).clear(); //For Youssef Waiting
      var source_range = Sheet3.getRange("A"+2+":AD"+2);
      var target_range = Sheet2.getRange("A"+(Sheet2.getLastRow()+1)+":AD"+(Sheet2.getLastRow()+1));
      source_range.moveTo(target_range);
     // Sheet3.getRange("A3:AD").moveTo(Sheet3.getRange("A2:AD"));
       Sheet3.deleteRow(2);
      Statistics();
    }
    }
    else
    {
      break;
    }
  }
   Statistics();
    if(flag==1)
   var name=Browser.msgBox('Done', 'Waiting insertion is done', Browser.Buttons.OK);
  
  }
}



//copying sumbimtted form responses
function onFormSubmit(e)
{
  var range=e.range;
  var row= range.getRow();
  var source_range = Sheet1.getRange("A"+(row)+":A"+(row));
   if(Sheet1.getRange(row, 5).getValue()==Exchange)
   {
   var last_row = Sheet5.getLastRow();
  var target_range = Sheet5.getRange("A"+(last_row+1)+":A"+(last_row+1));
   source_range.copyTo(target_range);
  source_range = Sheet1.getRange("B"+(row)+":Z"+(row));
  target_range = Sheet5.getRange("D"+(last_row+1)+":AD"+(last_row+1));
  source_range.copyTo(target_range);
  Sheet5.getRange(Sheet5.getLastRow(), 2).setValue(calculate(Sheet5,Sheet5.getLastRow()));
   }
   
   source_range = Sheet1.getRange("A"+(row)+":A"+(row))
   //Not waiting
  if((adults+children<max) && (Sheet3.getLastRow()==1))
  {
     
       var last_row = Sheet2.getLastRow();
       var target_range = Sheet2.getRange("A"+(last_row+1)+":A"+(last_row+1));
       source_range.copyTo(target_range);
       source_range = Sheet1.getRange("B"+(row)+":Z"+(row));
       target_range = Sheet2.getRange("D"+(last_row+1)+":AD"+(last_row+1));
       source_range.copyTo(target_range);
       Sheet2.getRange(Sheet2.getLastRow(), 2).setValue(calculate(Sheet2,Sheet2.getLastRow()));  //adding expected value
     
  }
  
  //In case of waiting
  else
  {
    var last_row = Sheet3.getLastRow();
    var target_range = Sheet3.getRange("A"+(last_row+1)+":A"+(last_row+1));
    source_range.copyTo(target_range);
    source_range = Sheet1.getRange("B"+(row)+":Z"+(row));
    target_range = Sheet3.getRange("D"+(last_row+1)+":AD"+(last_row+1));
    source_range.copyTo(target_range);
    Sheet3.getRange(Sheet3.getLastRow(), 2).setValue(calculate(Sheet3,Sheet3.getLastRow()));  //adding expected value
  }
 
 //Statisitcs edit
  Statistics();
   
}


//counting commas
function commasNumber (str)
{
  return (str.match(/,/g) || []).length;

}


//Check if it is car
function isCar(str){
if(str=="سيارة")
return true;
else
return false;
}

//count children factor in a certain row
function child(Sheet,row)
{
  var factor=0;
  var k=20;

  while(k<27)
  {
  var string=Sheet.getRange(row,k).getValue();
  if(string==half)
  {
    factor+=0.5;
    }
    else if (string==full)
    {
    factor+=1;
    }
    k+=3;
   }
    return factor;
}


function isFull(Sheet,row)
{
  var Counter=0;
  var z=20;

  while(z<27)
  {
  var string=Sheet.getRange(row,z).getValue();
  if(string==full)
  {
    Counter++;
    }
  
    z+=3;
   }
    return Counter;
}


//To find base value
function findTime(Sheet,row)
{
  var date=Sheet.getRange(row,1).getValue();
  var day=date.getDate();
  var month=date.getMonth()+1;
  if(month==11)
  {
    return 0;
  }
  if(month==12)
  {
    if(day<10)
    {
    return 0;
    }
    else if(day>=10&&day<24)
    {
      return 30;
    }
    else if(day>=24)
    return 60;
  }
  if(month==1)
  {
    if(day<7)
    {
      return 60;
    }
    if(day>=7&&day<21)
    {
      return 90;
    }
    if(day>=21)
    {
      return 120;
    }
  }
  
}

//caculate expected amount
function calculate (Sheet,row)
{
  var base=330+findTime(Sheet,row);
  var Days=commasNumber(Sheet.getRange(row,11).getValue());
  var go =Sheet.getRange(row, 9).getValue();
  var Come=Sheet.getRange(row, 10).getValue();
  return (1+child(Sheet,row))*(base+((Days)-4)*40)-50*(isCar(go)&&isCar(Come))*(1+isFull(Sheet,row));
}


//Calculate paid amount
function paid()
{
  var sum=0;
  for(var i=2;i<=Sheet2.getLastRow();i++){
    sum+=parseInt(Sheet2.getRange(i, 3).getValue()) || 0;
  }
  for(i=2;i<=Sheet3.getLastRow();i++){
    sum+=parseInt(Sheet3.getRange(i, 3).getValue()) || 0;
  }
  return sum;
}




//Count Full Children
function fullChildrenCount(Sheet){
var count=0;
for (var i=2;i<=Sheet2.getLastRow();i++)
  {
    if(Sheet.getRange(i, 20).getValue()==full)
    {
      count++;
    }
     if(Sheet.getRange(i, 23).getValue()==full)
    {
      count++;
    }
     if(Sheet.getRange(i, 26).getValue()==full)
    {
      count++;
    }
  }
  return count;

}


//Count half Children
function halfChildrenCount(Sheet){
var count=0;
for (var i=2;i<=Sheet.getLastRow();i++)
  {
    if(Sheet.getRange(i, 20).getValue()==half)
    {
      count++;
    }
     if(Sheet.getRange(i, 23).getValue()==half)
    {
      count++;
    }
     if(Sheet.getRange(i, 26).getValue()==half)
    {
      count++;
    }
  }
  return count;

}


//For the Statistics Sheet
function Statistics()
{
    max=parseInt(Sheet4.getRange(1, 3).getValue());
    children=parseInt(Sheet4.getRange(2, 2).getValue());
   adults=parseInt(Sheet4.getRange(6, 2).getValue()); 
}


//Count paid people
function paidCount()
{
  var count=0;
  for(var i=2;i<=Sheet2.getLastRow();i++){
   if(Sheet2.getRange(i, 3).getValue()!='')
     count++;
  }
  for(i=2;i<=Sheet3.getLastRow();i++){
     if(Sheet3.getRange(i, 3).getValue()!='')
         count++;
  }
  return count;
}


function stat (){
  Sheet4.getRange(1, 2).setFormula("Sum('IN'!C2:C)+Sum('Waiting' ! C2:C)");  //Total Amount Paid
  Sheet4.getRange(6, 2).setFormula("=Counta('IN'!A2:A)");      //Adults IN
  Sheet4.getRange(7, 2).setFormula("=Counta('IN'!C2:C)+Counta('Waiting'!C2:C)");     //Total Paid
  Sheet4.getRange(8, 2).setValue("=Counta('Waiting'!A2:A)");      //waiting
  
  var string="=countif('IN'!T2:T,"+'"'+full.toString()+'"'+")"+"+countif('IN'!W2:W,"+'"'+full.toString()+'"'+")"+"+countif('IN'!Z2:Z,"+'"'+full.toString()+'"'+")";
  Sheet4.getRange(2, 2).setFormula(string.toString());     //Full Children IN
    string="=countif('IN'!T2:T,"+'"'+half.toString()+'"'+")"+"+countif('IN'!W2:W,"+'"'+half.toString()+'"'+")"+"+countif('IN'!Z2:Z,"+'"'+half.toString()+'"'+")";
  Sheet4.getRange(3, 2).setFormula(string.toString());     //Half Children IN
    string="=countif('Waiting'!T2:T,"+'"'+full.toString()+'"'+")"+"+countif('Waiting'!W2:W,"+'"'+full.toString()+'"'+")"+"+countif('Waiting'!Z2:Z,"+'"'+full.toString()+'"'+")";
  Sheet4.getRange(4, 2).setFormula(string.toString());     //Full Children Waiting
      string="=countif('Waiting'!T2:T,"+'"'+half.toString()+'"'+")"+"+countif('Waiting'!W2:W,"+'"'+half.toString()+'"'+")"+"+countif('Waiting'!Z2:Z,"+'"'+half.toString()+'"'+")";
  Sheet4.getRange(5, 2).setFormula(string.toString());     //half children waiting
 
}
