import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

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
});

async function readCustomDocumentProperties() {
  await Word.run(async (context) => {
    var isDocuzenDoc=false;
    const properties = context.document.properties.customProperties;
    properties.load("key,value");
  
    await context.sync();
    const docProp= new DocuzenProperties();
    for (let i = 0; i < properties.items.length; i++)
      if(properties.items[i].key=="DVId"){
        isDocuzenDoc =true;
        docProp.dvid=properties.items[i].value;
      }
      if(properties.items[i].key=="SToken"){
        docProp.dvid=properties.items[i].value;
      }
      if(properties.items[i].key=="Uid"){
        docProp.dvid=properties.items[i].value;
      }
      if(properties.items[i].key=="logou"){
        docProp.dvid=properties.items[i].value;
      }

      //console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
  });
}



/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
