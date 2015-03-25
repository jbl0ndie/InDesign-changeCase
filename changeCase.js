//DESCRIPTION: Update of Dave Saunder's ChangeCaseOfSelectedStyle script for CS5 and later
//jbl0ndie: Fork of seddybell's update so that you can select by Paragraph Style instead of Character Style.
 
if ((app.documents.length != 0) && (app.selection.length != 0)) {
myDoc = app.activeDocument;
myStyles = myDoc.paragraphStyles; //jbl0ndie: replaced character with paragraph
myStringList = myStyles.everyItem().name;
myCaseList = ["UPPERCASE","lowercase", "Title Case", "Sentence case"]; //jbl0ndie: Changed myCaseList so it's consistent with InDesign's built in Change Case tool
myCases = [ChangecaseMode.uppercase, ChangecaseMode.lowercase, ChangecaseMode.titlecase, ChangecaseMode.sentencecase];
 
var myDialog = app.dialogs.add({name:"Case Changer"})
with(myDialog){
  with(dialogColumns.add()){
   with (dialogRows.add()) {
    with (dialogColumns.add()) {
     staticTexts.add({staticLabel:"Paragraph Style:"}); //jbl0ndie: replaced Character with Paragraph
    }
    with (dialogColumns.add()) {
     myStyle = dropdowns.add({stringList:myStringList,selectedIndex:0,minWidth:133}) ;
    }
   }
   with (dialogRows.add()) {
    with (dialogColumns.add()) {
     staticTexts.add({staticLabel:"Change Case to:"});
    }
    with (dialogColumns.add()) {
     myCase = dropdowns.add({stringList:myCaseList,selectedIndex:0,minWidth:133});
    }
   }
  }
}
var myResult = myDialog.show();
if (myResult != true){
  // user clicked Cancel
  myDialog.destroy();
  errorExit();
}
  theStyle = myStyle.selectedIndex;
  theCase = myCase.selectedIndex;
  myDialog.destroy();
 
 
  app.findTextPreferences = NothingEnum.NOTHING;
  app.changeTextPreferences = NothingEnum.NOTHING;
  app.findTextPreferences.appliedParagraphStyle = myStyles [theStyle]; var myFinds = myDoc.findText(); //jbl0ndie: replaced Character with Paragraph
  myLim = myFinds.length;
  for (var j=0; myLim > j; j++) {
   myFinds[j].texts[0].changecase(myCases[theCase]);
  }
 
 
} else {
errorExit();
}
 
 
// +++++++ Functions Start Here +++++++++++++++++++++++
 
 
function errorExit(message) {
if (arguments.length > 0) {
  if (app.version != 3) { beep() }
  alert(message);
}
exit();
}
