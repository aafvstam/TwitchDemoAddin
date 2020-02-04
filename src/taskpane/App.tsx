import * as React from "react";

import Progress from "./components/Progress";
import Header from "./components/common/Header";

import HomePage from "./pages/HomePage";
import AboutPage from "./pages/AboutPage";
import WatermarkPage from "./pages/WatermarkPage";

import { HashRouter as Router, Route, Switch } from "react-router-dom";
import { HeroListItem } from "./components/HeroList";

/* global Button Header, HeroList, HeroListItem, Progress, Word */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }
  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="container-fluid">
        <Header logo="assets/profile300x300.png" title={this.props.title} message="Twitch Demo 2020" />
        <Router>
          <div>
            <Switch>
              <Route exact path="/" component={HomePage} />
              <Route path="/Watermark" component={WatermarkPage} />
              <Route path="/About" component={AboutPage} />
            </Switch>
          </div>
        </Router>
      </div>
    );
  }
}
