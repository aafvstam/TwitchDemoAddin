import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./components/Header";
import HeroList, { HeroListItem } from "./components/HeroList";
//import Progress from "./components/Progress";
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
      const paragraph = context.document.body.insertParagraph("Hello About", Word.InsertLocation.end);

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
      /*
      return (
        <Progress
          title={title}
          logo="assets/profile300x300.png"
          message="Please sideload your addin to see app body."
        />
        );
       */
    }
    return (
      <div className="ms-welcome">
        <Header logo="assets/profile300x300.png" title={this.props.title} message="About" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
