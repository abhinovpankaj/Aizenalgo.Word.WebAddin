//import SubmitDocumentService from '../services/docuzenService' 

const AUTHENTICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcAuthentication";
const VERIFICATIONBASEURL = "https://demo.aizenalgo.com:9016/api/WordProc/WordProcSessionDetails";


let docProp= {};

Office.onReady(() => {
  // If needed, Office.js is ready to be called
 //readCustomDocumentProperties();
  
});

// Create a function for writing to the status div.
function updateStatus(message) {
  // var statusInfo = $('#status');
  // statusInfo[0].innerHTML += message + "<br/>";
  console.log(message);
}

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


//services
function SubmitDocumentService({stoken,dvid,uploadFile,fileName},type) {

  const endpoint = `${VERIFICATIONBASEURL}?SessionId=${stoken}&DocID=${dvid}&Mode=${type}`;
  
  //create file
  sendFile(endpoint);
               
}
var fileData;
// Get all of the content from  Word document in 100-KB chunks of text.
function sendFile(endpoint) {
  Office.context.document.getFileAsync("compressed",
      { sliceSize: 4000000 },
      function (result) {

          if (result.status == Office.AsyncResultStatus.Succeeded) {

              // Get the File object from the result.
              var myFile = result.value;
              var state = {
                  file: myFile,
                  counter: 0,
                  sliceCount: myFile.sliceCount
              };

              updateStatus("Getting file of " + myFile.size + " bytes");
              getSlice(state,endpoint);
             
          }
          else {              
              updateStatus(result.status);
              //return Promise.reject("failure");
          }
      });
}
// Get a slice from the file and then call sendSlice.
function getSlice(state,endpoint) {
  state.file.getSliceAsync(state.counter, function (result) {
      if (result.status == Office.AsyncResultStatus.Succeeded) {
          updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);                   
          sendSlice(result.value, state,endpoint);
          
      }
      else {
        
        updateStatus(result.status);
        return Promise.reject("failure");
      }
  });
}
function sendSlice(slice, state,endpoint) {
  var data = slice.data;

  if (data) {
    var file = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
    var formdata = new FormData();
    formdata.append("file", file);
    fetch(endpoint, {
      method: 'POST',
      body: formdata
    })
    .then(response => {
      if (!response.ok) throw (`invalid response: ${response.status}`);
          return response.json()
      })
    .then(data => console.log(data))
    .catch((err) => {
        console.log(err);
      })
      .finally(()=>{
        closeFile(state);
      });  
  }
  
}
function closeFile(state) {
  // Close the file when you're done with it.
  state.file.closeAsync(function (result) {

      // If the result returns as a success, the
      // file has been successfully closed.
      if (result.status == "succeeded") {        
          updateStatus("File closed.");
      }
      else {        
        updateStatus("File couldn't be closed.");
      }
  });
}

// The add-in command functions need to be available in global scope
g.submitDocument = submitDocument;
g.saveDocument = saveDocument;
