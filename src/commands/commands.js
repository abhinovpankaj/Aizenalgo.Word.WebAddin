import SubmitDocumentService from '../services/docuzenService' 

let docProp= {};
Office.onReady(() => {
  // If needed, Office.js is ready to be called
 //readCustomDocumentProperties();
  
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function submitDocument(event) {
  try {
      console.log("Inside submitDocument function");
      readCustomDocumentProperties();
      //console.log(Office.context.document.url);
    
      
  } catch (error) {
    console.log(error);
  }

  // const message = {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: "Submitting Docuzen document.",
  //   icon: "Icon.80x80",
  //   persistent: false,
  // };

  // Show a notification message
  //Office.context.mailbox.item.notificationMessages.replaceAsync("submitDocument", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function saveDocument(event) {
  console.log("Inside save function");
 
  event.completed();
}
function readCustomDocumentProperties() {
  console.log("Inside readcustom function,Commands.js");
  
   Word.run(async (context) => {
    //var isDocuzenDoc=false;
    const properties = context.document.properties.customProperties;
    properties.load("key,value");
  
    await context.sync();
    try{    

      for (let i = 0; i < properties.items.length; i++){        
        if(properties.items[i].key=="DVId"){
          //isDocuzenDoc =true;
          docProp.dvid=properties.items[i].value;
        }
        if(properties.items[i].key=="SToken"){
          docProp.stoken=properties.items[i].value;
        }
        if(properties.items[i].key=="Uid"){
          docProp.uid=properties.items[i].value;
        }
        if(properties.items[i].key=="logou"){
          docProp.logou=properties.items[i].value;
        }
      }
      //set document name and path
      var uploadFilePath = Office.context.document.url;
      var pieces = uploadFilePath.split('\\');
      var filename = pieces[pieces.length-1];
      docProp.fileName=  filename;
      docProp.uploadFile = uploadFilePath;
      console.log(docProp) ;    
      
      SubmitDocumentService(docProp,1);
    
    }
    catch(error){
      console.log("read doc property:" + error.stack);
    }
    
  });
}
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.submitDocument = submitDocument;
g.saveDocument = saveDocument;
