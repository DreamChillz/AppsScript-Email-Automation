function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName('Email Templates');
  const templatesData = templatesSheet.getDataRange().getValues();
  
  const templates = [];
  for (let i = 1; i < templatesData.length; i++) {
    templates.push(templatesData[i][0]);
  }
  
  return templates;
}

function getTemplatePlaceholders(templateName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName('Email Templates');
  const templatesData = templatesSheet.getDataRange().getValues();
  
  const fixedValuesSheet = ss.getSheetByName('Fixed Values');
  const fixedValuesData = fixedValuesSheet.getDataRange().getValues();
  const fixedValuesSet = new Set();
  for (let i = 1; i < fixedValuesData.length; i++) {
    fixedValuesSet.add(fixedValuesData[i][0]);
  }

  let templateBody = '';
  for (let i = 1; i < templatesData.length; i++) {
    if (templatesData[i][0] === templateName) {
      templateBody = templatesData[i][2]; // Assuming the body is in the third column (index 2)
      break;
    }
  }
  
  const regex = /\[([^\]]+)\]/g;
  const placeholdersSet = new Set();
  let match;
  while ((match = regex.exec(templateBody)) !== null) {
    const placeholder = match[1];
    if (placeholder !== 'First Name' && !fixedValuesSet.has(placeholder)) {
      placeholdersSet.add(placeholder);
    }
  }
  
  return Array.from(placeholdersSet);
}

function generateEmail(templateName, email, placeholders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName('Email Templates');
  const templatesData = templatesSheet.getDataRange().getValues();
  
  const fixedValuesSheet = ss.getSheetByName('Fixed Values');
  const fixedValuesData = fixedValuesSheet.getDataRange().getValues();
  const fixedValuesMap = {};
  for (let i = 1; i < fixedValuesData.length; i++) {
    fixedValuesMap[fixedValuesData[i][0]] = fixedValuesData[i][1];
  }

  let template = null;
  for (let i = 1; i < templatesData.length; i++) {
    if (templatesData[i][0] === templateName) {
      template = {
        subject: templatesData[i][1],
        body: templatesData[i][2]
      };
      break;
    }
  }
  
  if (!template) {
    throw new Error('Template not found');
  }

  const recipientName = getFirstNameByEmail(email);
  if (!recipientName) {
    throw new Error('Customer not found for the given email');
  }
  
  let subject = template.subject.replace(/\[First Name\]/g, recipientName);
  let body = template.body.replace(/\[First Name\]/g, recipientName);

  for (const [key, value] of Object.entries(placeholders)) {
    const regex = new RegExp(`\\[${key}\\]`, 'g');
    subject = subject.replace(regex, value);
    body = body.replace(regex, value);
  }

  for (const [key, value] of Object.entries(fixedValuesMap)) {
    const regex = new RegExp(`\\[${key}\\]`, 'g');
    subject = subject.replace(regex, value);
    body = body.replace(regex, value);
  }
  
  return {
    subject: subject,
    body: body
  };
}

function getFirstNameByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customerDataSheet = ss.getSheetByName('Customer Data');
  const customerData = customerDataSheet.getDataRange().getValues();
  
  for (let i = 1; i < customerData.length; i++) {
    if (customerData[i][1] === email) { // Assuming email is in the second column (index 1)
      return customerData[i][0]; // Assuming name is in the first column (index 0)
    }
  }
  
  return null;
}

function importEmailsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customerDataSheet = ss.getSheetByName('Customer Data');
  const customerData = customerDataSheet.getDataRange().getValues();
  const emails = [];
  
  for (let i = 1; i < customerData.length; i++) {
    if (customerData[i][1]) {
      emails.push(customerData[i][1]);
    }
  }
  
  return emails;
}

function generateEmailForMultipleRecipients(templateName, emails, placeholders) {
  const generatedEmails = emails.map(email => {
    try {
      return {
        email: email,
        ...generateEmail(templateName, email, placeholders)
      };
    } catch (error) {
      return {
        email: email,
        error: error.message
      };
    }
  });
  return generatedEmails;
}

function sendEmails(emails) {
  const results = emails.map(email => {
    try {
      GmailApp.sendEmail(email.email, email.subject, email.body);
      return { email: email.email, status: "Success" };
    } catch (error) {
      return { email: email.email, status: "Failed", error: error.message };
    }
  });

  return results;
}




function getTemplateBody(templateName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName('Email Templates');
  const templatesData = templatesSheet.getDataRange().getValues();

  let templateBody = '';
  for (let i = 1; i < templatesData.length; i++) {
    if (templatesData[i][0] === templateName) {
      templateBody = templatesData[i][2]; // Assuming the body is in the third column (index 2)
      break;
    }
  }

  return templateBody;
}

