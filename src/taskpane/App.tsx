import * as React from "react";

import Progress from "./components/Progress";
import Header from "./components/common/Header";
import HomePage from "./pages/HomePage";
import WatermarkPage from "./pages/WatermarkPage";

import { Route, Switch } from "react-router-dom";
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

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      let hostInfoResult = this.getHostInfo();

      // insert a paragraph at the end of the document.
      const paragraphHostInfo = context.document.body.insertParagraph(hostInfoResult, Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraphHostInfo.font.color = "Red";

      await context.sync();
    });
  };

  getHostInfo() {
    var _requirements = Office.context.requirements;
    var types = ["Excel", "Word"];
    var minVersions = ["Preview", "1.6", "1.5", "1.4", "1.3", "1.2", "1.1", "1.0"]; // Start with the highest version

    // Loop through types and minVersions
    for (var type in types) {
      for (var minVersion in minVersions) {
        // Append "Api" to the type for set name, i.e. "ExcelApi" or "WordApi"
        if (_requirements.isSetSupported(types[type] + "Api", minVersions[minVersion])) {
          return minVersions[minVersion];
        }
      }
    }

    return "Nothing";
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
        <Header logo="assets/twitch_80x80.png" title={this.props.title} message="Twitch Demo Add-in" />
        <Switch>
          <Route path="/" exact component={HomePage} />
          <Route path="/" exact component={WatermarkPage} />
        </Switch>
      </div>
    );
  }
}
