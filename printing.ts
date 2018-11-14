// const printSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
//   let printerId = "616c0920-968d-1759-9276-dc716a41e2af" 
//   var title = sheet.getName();
//   var ticket = {
//     version: "1.0",
//     print: {
//       color: {
//         type: "STANDARD_COLOR",
//         vendor_id: "Color"
//       },
//       duplex: {
//         type: "NO_DUPLEX"
//       }
//     }
//   }
//   var content = sheet
//   var optContentType = "application/pdf";
//   var optTag = "";
      
//   return submit(printerId, title, ticket, content, optContentType, optTag)
// }

// function submit(printerid, title, ticket, content, optContentType, optTag) {
  
  
//   var contentType = optContentType || "";
//   var tag = optTag || "";
  
//   var params = {
//     service: "submit",
//     printerid: printerid,
//     title: title,
//     ticket: JSON.stringify(ticket),
//     content: content,
//     contentType: contentType,
//     tag: tag
//   }
    
//   return doGetPattern_({}, constructConsentScreen_, restCallGCP_, null, params);
// }

// /*Google Cloud Print REST Connection Call for all Service Interfaces
// */
// function restCallGCP_(accessToken, params) {
//   var service = params.service;
      
//   var options = {
//      method: "POST",
//      headers: {
//        authorization: "Bearer " + accessToken,
//      },
//      payload: params
//    };  
//   var result = UrlFetchApp.fetch("https://www.google.com/cloudprint/"+service, options);
//   return result.getContentText();
// }

// function doGetPattern_(e, consentScreen, theWork,optPackageName, optParams) {
//   // set up authentication
//   var packageName = optPackageName || '';
//   var authenticationPackage = getAuthenticationPackage_ (packageName);
//   if (!authenticationPackage) {
//     throw "You need to set up your credentials one time";
//   }

//   var eo = new EzyOauth2 ( authenticationPackage, "getAccessTokenCallback", undefined, {work:theWork.name,package_name:packageName} );
  
//   // eo will have checked for an unexpired access code, or got a new one with a refresh code if it was possible, and we'll already have it
//   if (eo.isOk()) {
//     // should save the updated properties for next time
//     setAuthenticationPackage_ (authenticationPackage);
//     // good to do whatever we're here to do
//     var params = optParams || {};
//     return theWork (eo.getAccessToken(), params);
//   }
  
//   else {

//     // start off the oauth2 dance - you'll want to pretty this up probably
//       return HtmlService.createHtmlOutput ( consentScreen(eo.getUserConsentUrl()) );
//   }
// }

// function constructConsentScreen_ (consentUrl) {
//   return '<a href = "' + consentUrl + '">Authenticate</a> ';
// }