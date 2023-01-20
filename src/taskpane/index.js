import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { exit } from "process";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Docuzen Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;

  //get the Document Properties.
  readCustomDocumentProperties();
  render(App);
  // Office.context.ui.displayDialogAsync('https://localhost:3000/taskpane.html', {height: 60, width: 40});

});

async function readCustomDocumentProperties() {
  
  await Word.run(async (context) => {
    var isDocuzenDoc=false;
    
    const properties = context.document.properties.customProperties;
    properties.load("key,value");
  
    await context.sync();
    try{
      
      for (let i = 0; i < properties.items.length; i++){
        //console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
        if(properties.items[i].key=="DVId"){
          isDocuzenDoc =true;
          break;
        }
      }
      
      //hideAizenAlgoGroup(isDocuzenDoc);  
      
    }
    catch(error){
      console.log("read doc property:" + error.stack);
    }
    
  });
}

function hideAizenAlgoGroup(state) {
  try {
    console.log("Inside hide tab function");
    var parentGroup = {
      id: "Aizenalgo.CommandsGroup",
      visible:state
    };
    var parentTab = {
      id: "Docuzen.Tab1",
      groups: [parentGroup]
    };
    var ribbonUpdater = { tabs: [parentTab] };
    Office.ribbon.requestUpdate(ribbonUpdater);
    
    console.log(state);
  } catch (error) {
    console.log("Hide Aizenalgo Group:" + error.stack);
  }
}

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
