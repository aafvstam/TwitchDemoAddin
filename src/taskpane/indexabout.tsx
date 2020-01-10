import "office-ui-fabric-react/dist/css/fabric.min.css";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import AppAbout from "./AppAbout";

/* global AppContainer, Component, document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Maarten's About Task Pane";

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(AppAbout);
};

/* Initial render showing a progress bar */
render(AppAbout);

if ((module as any).hot) {
  (module as any).hot.accept("./App", () => {
    const NextApp = require("./App").default;
    render(NextApp);
  });
}
