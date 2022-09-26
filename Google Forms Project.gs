// @ts-nocheck
/**
 * 
 * Introduction to Apps Script Project
 * Project: Google Form Coding Questionnaire
 * 
 */

// talk to the Google Form
// ID of the Form
// https://docs.google.com/forms/d/1T_EKSDjoFeh9ZmN8-zZRHPdgPcCs0g0wIIs98w7une4/edit

const FORM_ID = '1T_EKSDjoFeh9ZmN8-zZRHPdgPcCs0g0wIIs98w7une4';

// add a custom menu to our Sheet
function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Questionnaire Menu')
    .addItem('Update Form', 'updateForm_v2')
    .addItem('Send Emails', 'sendEmail_v2')
    .addToUi();
}

// get the IDs of the Form components
function getFormIDs() {

  const form = FormApp.openById(FORM_ID);
  const formItems = form.getItems(); // array of form items

  // loop over array
  // print out form items title & ID
  formItems.forEach(item => console.log(item.getTitle() + ' ' +  item.getId()));

}

/*
Name 862364853
Email Address 1441315353
Do you have any prior experience with coding? 1891947248
What programming languages do you use? 150791220
*/

// update the form question from Sheet

// Version 2
function updateForm_v2() {

  // get list of languages in setup Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName('setup');
  const setupVals = setupSheet.getRange(2,1,setupSheet.getLastRow()-1,1).getValues().flat();
  // console.log(setupVals);
  // [ 'None', 'App Script' ]

  // get list of languages submitted via the Form
  const responseSheet = ss.getSheetByName('Form Responses 1');
  const data = responseSheet.getRange(2,5,responseSheet.getLastRow()-1,1).getValues().flat();
  const submitVals = data.join().split(',');
  //console.log(data);
  //console.log(submitVals);
  // [ 'None','Apps Script','Apps Script',' JavaScript','Apps Script',' Python' ]

  // get list of languages in Form question
  const form = FormApp.openById(FORM_ID);
  const formCheckboxChoices = form.getItemById('150791220').asCheckboxItem().getChoices();
  const formCheckboxValues = formCheckboxChoices.map(x => x.getValue());
  //console.log(formCheckboxValues);

  // consolidate list of languages
  const allLangs = [...formCheckboxValues,...setupVals,...submitVals];
  //console.log(allLangs);
  /*
  [ 'None',
  'Apps Script',
  'JavaScript',
  'Python',
  'None',
  'Apps Script',
  'JavaScript',
  'Python',
  '',
  ' ',
  ' ',
  ' ' ]
  */

  // remove leading and trailing spaces from languages
  const trimLangs = allLangs.map(item => item.trim());
  //console.log(trimLangs);
  /*
  [ 'None',
  'Apps Script',
  'JavaScript',
  'Python',
  'None',
  'Apps Script',
  'JavaScript',
  'Python',
  '',
  '',
  '',
  '' ]
  */
 
  // sort list of languages
  trimLangs.sort();
  //console.log(trimLangs);
  /*
  [ '',
  '',
  '',
  '',
  'Apps Script',
  'Apps Script',
  'JavaScript',
  'JavaScript',
  'None',
  'None',
  'Python',
  'Python' ]
  */
  
  // dedup list of languages
  let finalLangList = trimLangs.filter((lang,i) => trimLangs.indexOf(lang) === i);
  // console.log(finalLangList);
  // [ '', 'Apps Script', 'JavaScript', 'None', 'Python' ]

  // remove any blanks
  finalLangList = finalLangList.filter(item => item !== 'None');
  // console.log(finalLangList);
  // [ 'Apps Script', 'JavaScript', 'None', 'Python' ]

  // move 'None' to front of array
  finalLangList = finalLangList.filter(item => item.length !== 0);
  // console.log(finalLangList);
  // [ 'Apps Script', 'JavaScript', 'Python' ]
  finalLangList.unshift('None');
  // console.log(finalLangList);
  // [ 'None', 'Apps Script', 'JavaScript', 'Python' ]

  // turn into double array notation for Sheets
  const finalDoubleArray = finalLangList.map(lang => [lang]);
  // console.log(finalDoubleArray);
  // [ [ 'None' ], [ 'Apps Script' ], [ 'JavaScript' ], [ 'Python' ] ]

  // paste into setup Sheet 
  setupSheet.getRange('A2:A').clear();
  setupSheet.getRange(2,1,finalLangList.length,1).setValues(finalDoubleArray);

  // copy into the Form
  form.getItemById('150791220').asCheckboxItem().setChoiceValues(finalLangList);

}

// Version 1
function updateForm_v1() {

  // get list of languages from Google Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName('setup');

  const langVals = setupSheet.getRange(2,1,setupSheet.getLastRow()-1,1).getValues();
  console.log(langVals);
  // [ [ 'None' ], [ 'Apps Script' ] ]

  const langValsFlat = langVals.map(item => item[0]); // ['None'] => 'None'
  console.log(langValsFlat);
  // [ 'None', 'Apps Script' ] 


  // get hold of the Form and the question
  const form = FormApp.openById(FORM_ID);
  const langsCheckboxQuestion = form.getItemById('150791220').asCheckboxItem();

  // populate the form question with the language list
  // array of strings
  // ['Dogs','Cats']
  langsCheckboxQuestion.setChoiceValues(langValsFlat);

}


// automatically send emails to respondents with their information

// Version 2
function sendEmail_v2() {

  //get the spreadsheet information
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('Form Responses 1');
  const data = responseSheet.getDataRange().getValues();
  //console.log(data);

  // remove the header
  data.shift();
  //console.log(data);

  // loop over the rows
  data.forEach((row,i) => {

    // identify the ones I haven't replied to
    if(row[5] === '') { 
      
      // get the email address
      const name = row[1];
      const email = row[2];
      const answer = row[3]; // yes or no
      const langs = row[4]; // list of languages

      // write the email
      const subject = 'Thank you responding to the Apps Script questionnaire!';
      let body = '' + i;
      // console.log(body);

      // change the body for yes and no
      // yes answer
      if (row [3] === 'Yes') {
        body = 'Hi ' + name + `,<br><br>
          Thank you for responding to our 2022 Developer Survey!<br><br>
          Your feedback is greatly appreciated!<br><br>
          You reported experience with the following coding languages:<br><br>
          <em>` + langs + `</em><br><br>
          Thanks,<br>
          Keola`;
      }
      // no answer
      else {
        body = 'Hi ' + name + `,<br><br>
         Thank you for responding to our 2022 Developer Survey!<br><br>
         Your feedback is greatly appreciated!<br><br>
         You reported not having any experience with coding, so here's a resource to get started:<br><br>
         <a href="https://www.benlcollins.com/spreadsheets/starting-gas/">Getting started with Apps Script</a><br><br>         
         Thanks,<br>
         Keola`;
      }
     console.log(email);
     console.log(subject);
     console.log(body);

     // send the email
     GmailApp.sendEmail(email,subject,'',{htmlBody: body});

     // mark as sent
     const d = new Date();
     responseSheet.getRange(i + 2,6).setValue(d);

    }
    else {
      console.log('No email sent for this row');
    }

  });

}

// Version 1
function sendEmail_v1() {

  //get the spreadsheet information
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('Form Responses 1');
  const data = responseSheet.getDataRange().getValues();
  //console.log(data);

  // remove the header
  data.shift();
  //console.log(data);

  // loop over the rows
  data.forEach((row,i) => {

    // identify the ones I haven't replied to
    if(row[5] === '') { 
      
      // get the email address
      const email = row[2];
      console.log(email);

      // write the email
      const subject = 'Thank you responding to the Apps Script questionnaire!';
      let body = '' + i;
      // console.log(body);

      // change the body for yes and no
      // yes answer
      if (row [3] === 'Yes') {
        body = 'TBC - YES answer'
      }
      // no answer
      else {
        body = 'TBC - NO answer'
      }

      // send the email
      GmailApp.sendEmail(email,subject,body);

      // mark as sent
      const d = new Date();
      responseSheet.getRange(i + 2,6).setValue(d);

    }
    else {
      console.log('No email sent for this row');
    }

  });
}

