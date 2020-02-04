import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { BrowserRouter } from "react-router-dom";
/* global AppContainer, Component, document, Office, module, require */

initializeIcons();

// Set to true if you want to run the addin in a common browser (not in an Office Host like MSWord)
let isOfficeInitialized = true;

const title = "Twitch Demo";

const render = Component => {
  ReactDOM.render(
    <BrowserRouter>
      <AppContainer>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </AppContainer>
    </BrowserRouter>,
    document.getElementById("container")
  );
};


/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
