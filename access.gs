function LocalRefresh() {
  const doc = DocumentApp.getActiveDocument(); // current doc
  AutomatedImageRefresher.onOpen(doc, 2, '', false); 
  
  /* onOpen(doc, startNum, endText, footer) where startNum is the starting row number, 
    endText is the alt text AFTER the last one you want to check (format: 'alttext'),
    footer is if you want it to check the footer as well (set to true if you do)
  You can change startNum,endText and footer to improve efficiency
  Default is start on row 2 and go till the alt text cell is empty with footer set to false*/
}
